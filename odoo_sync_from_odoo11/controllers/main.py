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
@http.route('/odoo_sync/purchase_order', type='json', auth='user', csrf=False, methods=['POST'])
def receive_purchase_data(self, **post):
        try:
            raw_data = request.httprequest.data.decode('utf-8')
            data = json.loads(raw_data) if raw_data else {}
            _logger.info("PurchaseOrder reçu : %s", json.dumps(data, indent=2))
            
        except Exception as e:
            _logger.exception("Erreur reception AccountInvoice : %s", e)
            return {"status": "error", "message": str(e)}