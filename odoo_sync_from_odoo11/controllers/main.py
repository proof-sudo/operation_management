from odoo import http, SUPERUSER_ID
from odoo.http import request
import logging
import json

_logger = logging.getLogger(__name__)

class OdooSyncController(http.Controller):

    @http.route('/odoo_sync/sale_order', type='json', auth='public', csrf=False, methods=['POST'])
    def receive_sale_order(self, **post):
        try:
            # Récupération du payload JSON
            raw_data = request.httprequest.data.decode('utf-8')
            data = json.loads(raw_data) if raw_data else {}
            _logger.info("SaleOrder reçu : %s", json.dumps(data, indent=2))

            # Vérif partner_id
            if not data.get("partner_id"):
                return {"status": "error", "message": "partner_id manquant dans la requête"}

            # Client
            partner_id, partner_name = data['partner_id']
            partner = request.env['res.partner'].sudo().search([('id', '=', partner_id)], limit=1)
            if not partner:
                partner = request.env['res.partner'].sudo().create({'name': partner_name})
                _logger.info("Client créé : %s", partner.name)

            # Entrepôt
            warehouse = False
            if data.get('warehouse_id'):
                warehouse_id, warehouse_name = data['warehouse_id']
                warehouse = request.env['stock.warehouse'].sudo().search([('id', '=', warehouse_id)], limit=1)
                if not warehouse:
                    warehouse = request.env['stock.warehouse'].sudo().create({
                        'name': warehouse_name,
                        'code': warehouse_name[:5].upper(),
                    })
                    _logger.info("Entrepôt créé : %s", warehouse.name)

            # Utilisateur
            user = None
            if data.get('user_id'):
                user_id, user_name = data['user_id']
                user = request.env['res.users'].sudo().search([('id', '=', user_id)], limit=1)
            if not user:
                user = request.env['res.users'].sudo().browse(SUPERUSER_ID)
                _logger.warning("Utilisateur non trouvé, utilisation admin : %s", user.login)

            # Vérifier si le SaleOrder existe déjà (évite les doublons)
            sale_order = request.env['sale.order'].sudo().search([('name', '=', data['name'])], limit=1)
            if sale_order:
                _logger.info("SaleOrder %s déjà existant, aucun doublon créé.", data['name'])
                return {"status": "success", "sale_order_id": sale_order.id}

            # Création du SaleOrder
            sale_order_vals = {
                'name': data['name'],
                'partner_id': partner.id,
                'user_id': user.id,
                'amount_total': data.get('amount_total', 0),
                'warehouse_id': warehouse.id if warehouse else False,
                'project_name': data.get('project', False),
            }
            sale_order = request.env['sale.order'].sudo().create(sale_order_vals)
            _logger.info("SaleOrder créé localement : %s", sale_order.name)

            # Création des lignes de commande
            for line in data.get('order_lines_data', []):
                # Toujours créer le produit
                product_id, product_name = line['product_id']
                product = request.env['product.product'].sudo().create({
                    'name': product_name,
                    'list_price': line.get('price_unit', 0),
                })
                _logger.info("Produit créé : %s", product.name)

                # Ligne de commande
                line_vals = {
                    'order_id': sale_order.id,
                    'product_id': product.id,
                    'product_uom_qty': line.get('product_uom_qty', 1),
                    'price_unit': line.get('price_unit', 0),
                    'name': line.get('name', 'Produit inconnu'),
                    'tax_id': line.get('taxes_id', []),
                }
                request.env['sale.order.line'].sudo().create(line_vals)

            return {"status": "success", "sale_order_id": sale_order.id}

        except Exception as e:
            _logger.exception("Erreur reception SaleOrder : %s", e)
            return {"status": "error", "message": str(e)}

    @http.route('/odoo_sync/account_invoice', type='json', auth='user', csrf=False, methods=['POST'])
    def receive_account_invoice(self, **post):
        try:
            raw_data = request.httprequest.data.decode('utf-8')
            data = json.loads(raw_data) if raw_data else {}
            _logger.info("AccountInvoice reçu : %s", json.dumps(data, indent=2))

            # Vérif partenaire
            if not data.get("partner_id"):
                return {"status": "error", "message": "partner_id manquant dans la requête"}

            partner = request.env['res.partner'].sudo().search([('id', '=', data['partner_id'][0])], limit=1)
            if not partner:
                partner = request.env['res.partner'].sudo().create({'name': data['partner_id'][1]})
                _logger.info("Client créé : %s", partner.name)

            # Utilisateur
            user = False
            if data.get('user_id'):
                user = request.env['res.users'].sudo().search([('id', '=', data['user_id'][0])], limit=1)
            if not user:
                user = request.env['res.users'].sudo().browse(SUPERUSER_ID)
                _logger.warning("Utilisateur non trouvé, utilisation admin : %s", user.login)

            # Création de la facture
            invoice_vals = {
                'move_type': 'out_invoice',
                'partner_id': partner.id,
                'invoice_date': data.get('date_invoice', None),
                'invoice_origin': data.get('origin', ''),
                'amount_total': data.get('amount_total', 0),
            }
            invoice = request.env['account.move'].sudo().create(invoice_vals)
            _logger.info("AccountInvoice créé localement : %s", invoice.name)

            return {"status": "success", "invoice_id": invoice.id}

        except Exception as e:
            _logger.exception("Erreur reception AccountInvoice : %s", e)
            return {"status": "error", "message": str(e)}
        
    @http.route('/odoo_sync/purchase_order', type='json', auth='public', csrf=False, methods=['POST'])
    def receive_purchase_data(self, **post):
        try:
            _logger.info("Début réception PurchaseOrder")
            raw_data = request.httprequest.data.decode('utf-8')
            data = json.loads(raw_data) if raw_data else {}
            _logger.info("PurchaseOrder reçu : %s", json.dumps(data, indent=2))

            # Traitement des données
            result = self._process_purchase_order(data)
            return result

        except Exception as e:
            _logger.exception("Erreur reception PurchaseOrder : %s", e)
            return {"status": "error", "message": str(e)}

    def _process_purchase_order(self, data):
        """Traite et crée la commande d'achat dans Odoo 18"""
        try:
            # Rechercher le fournisseur
            partner_id = self._find_partner(data.get('partner_id'))
            if not partner_id:
                return {"status": "error", "message": "Fournisseur non trouvé"}

            # Vérifier si la commande existe déjà
            existing_order = request.env['purchase.order'].sudo().search([
                ('name', '=', data.get('name'))
            ], limit=1)

            if existing_order:
                _logger.info("Commande existe déjà: %s", data.get('name'))
                return {"status": "success", "message": "Commande déjà existante", "purchase_id": existing_order.id}

            # Gérer le dossier
            dossier_data = data.get('dossier_data', {})
            dossier_name = self._extract_dossier_name(dossier_data)
            ref_fp = dossier_data.get('ref_bc_customer', '')
            client_info = dossier_data.get('client_id', False)
            client_id = False
            if client_info and isinstance(client_info, list):
                client_id = self._find_partner(client_info)

            # Préparer les valeurs pour la commande
            order_vals = {
                'name': data.get('name'),
                'partner_id': partner_id,
                'date_order': data.get('date_order'),
                'partner_ref': data.get('partner_ref', ''),
                'date_approve': data.get('date_approve'),
                'currency_id': self._find_currency(data.get('currency_id')),
                'notes': data.get('notes', ''),
                'origin': f"Sync Odoo11: {data.get('name')}",
                'company_id': request.env.company.id,
                'ref_Fp': ref_fp,
            }
            if client_id:
                order_vals['client_id'] = client_id

            # Ajouter partner_ref s'il existe
            if data.get('partner_ref'):
                order_vals['partner_ref'] = data.get('partner_ref')

            # Ajouter le dossier_id (nom du dossier)
            if dossier_name:
                order_vals['dossier_id'] = dossier_name

            # Créer la commande
            purchase_order = request.env['purchase.order'].sudo().create(order_vals)

            # Créer les lignes de commande
            order_lines_data = data.get('order_lines_data', [])
            for line_data in order_lines_data:
                self._create_order_line(purchase_order.id, line_data)

            # Confirmer la commande
            purchase_order.button_confirm()

            _logger.info("✅ Commande créée avec succès: %s (ID: %s, Dossier: %s)", purchase_order.name, purchase_order.id, dossier_name)

            return {
                "status": "success", 
                "message": "Commande créée avec succès", 
                "purchase_id": purchase_order.id,
                "purchase_name": purchase_order.name,
                "dossier_id": dossier_name
            }

        except Exception as e:
            _logger.exception("Erreur traitement PurchaseOrder: %s", str(e))
            return {"status": "error", "message": f"Erreur traitement: {str(e)}"}

    def _extract_dossier_name(self, dossier_data):
        """Extrait le nom du dossier depuis les données"""
        if not dossier_data:
            return False

        # Priorité: name, puis project_name
        dossier_name = dossier_data.get('name')
        if not dossier_name:
            dossier_name = dossier_data.get('project_name')

        return dossier_name

    def _find_partner(self, partner_data):
        """Trouve le fournisseur par nom"""
        if not partner_data:
            return False

        partner_name = partner_data[1] if isinstance(partner_data, list) else str(partner_data)

        # Rechercher par nom exact
        partner = request.env['res.partner'].sudo().search([
            ('name', '=ilike', partner_name)
        ], limit=1)

        if not partner:
            # Créer le fournisseur
            partner = request.env['res.partner'].sudo().create({
                'name': partner_name,
                'company_type': 'company',
                'supplier_rank': 1,
            })
            _logger.info("Nouveau fournisseur créé: %s", partner_name)

        return partner.id

    def _find_currency(self, currency_data):
        """Trouve la devise par nom"""
        if not currency_data:
            return request.env.company.currency_id.id

        currency_name = currency_data[1] if isinstance(currency_data, list) else str(currency_data)
        
        # Rechercher la devise
        currency = request.env['res.currency'].sudo().search([
            ('name', '=ilike', currency_name)
        ], limit=1)

        return currency.id if currency else request.env.company.currency_id.id

    def _create_order_line(self, order_id, line_data):
        """Crée une ligne de commande d'achat"""
        try:
            # Trouver ou créer le produit
            product_id = self._find_or_create_product(line_data.get('product_id'))

            # Préparer les valeurs de la ligne
            line_vals = {
                'order_id': order_id,
                'product_id': product_id,
                'product_qty': line_data.get('product_qty', 1.0),
                'price_unit': line_data.get('price_unit', 0.0),
                'name': line_data.get('name', ''),
                'date_planned': line_data.get('date_planned'),
            }

            # Créer la ligne
            order_line = request.env['purchase.order.line'].sudo().create(line_vals)

            # Gérer les taxes si disponibles
            taxes_data = line_data.get('taxes_id')
            if taxes_data and isinstance(taxes_data, list) and len(taxes_data) > 2:
                tax_ids = taxes_data[2]  # Récupérer les IDs de taxes
                if tax_ids:
                    # Chercher les taxes par nom (approximatif)
                    taxes = request.env['account.tax'].sudo().search([
                        ('type_tax_use', '=', 'purchase')
                    ], limit=1)
                    if taxes:
                        order_line.taxes_id = taxes

            return order_line.id

        except Exception as e:
            _logger.error("Erreur création ligne commande: %s", str(e))
            return False

    def _find_or_create_product(self, product_data):
        """Trouve ou crée un produit par nom"""
        if not product_data:
            # Retourner un produit générique si non spécifié
            generic_product = request.env['product.product'].sudo().search([
                ('default_code', '=', 'GENERIC')
            ], limit=1)
            
            if not generic_product:
                generic_product = request.env['product.product'].sudo().create({
                    'name': 'Produit Générique',
                    'default_code': 'GENERIC',
                    'type': 'service',
                    'purchase_ok': True,
                })
            return generic_product.id

        product_name = product_data[1] if isinstance(product_data, list) else str(product_data)

        # Rechercher par nom
        product = request.env['product.product'].sudo().search([
            ('name', '=ilike', product_name)
        ], limit=1)

        if not product:
            # Créer le produit
            product = request.env['product.product'].sudo().create({
                'name': product_name,
                'type': 'service',  # ou 'product' selon le besoin
                'purchase_ok': True,
                'sale_ok': False,
                'default_code': f"PROD_{product_name[:20]}",
            })
            _logger.info("Nouveau produit créé: %s", product_name)

        return product.id