from odoo import models, fields, api
import logging
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

class ProjectInvoiceRequest(models.Model):
    _name = 'project.invoice.request'
    _description = 'Demande de Facturation Projet'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _order = 'create_date desc'
    
    name = fields.Char(string='Référence', required=True, default='Nouveau', copy=False)
    project_id = fields.Many2one('project.project', string='Projet', required=True, tracking=True)
    sale_order_id = fields.Many2one('sale.order', string='Commande Client', required=True)
    line_ids = fields.One2many('project.invoice.request.line', 'request_id', string='Lignes')
    
    # NOUVEAUX CHAMPS AJOUTÉS
    description = fields.Text(string='Description / Notes')
    attachment_ids = fields.Many2many(
        'ir.attachment',
        'project_invoice_request_attachment_rel',
        'request_id', 
        'attachment_id',
        string='Documents Joints'
    )
    total_amount = fields.Monetary( string='Montant Total', compute='_compute_total_amount', store=True, currency_field='currency_id')
    document_count = fields.Integer(
        string='Nombre de Documents',
        compute='_compute_document_count'
    )
    
    state = fields.Selection([
        ('draft', 'Brouillon'),
        ('submitted', 'Soumise'),
        ('approved', 'Approuvée'),
        ('invoiced', 'Facturée'),
        ('rejected', 'Rejetée'),
    ], string='État', default='draft', tracking=True)
    
   
    currency_id = fields.Many2one(
        related='sale_order_id.currency_id', 
        store=True
    )
    invoice_id = fields.Many2one('account.move', string='Facture Créée', readonly=True)
    
    @api.depends('attachment_ids')
    def _compute_document_count(self):
        """Calcule le nombre de documents joints"""
        for request in self:
            request.document_count = len(request.attachment_ids)
            
    @api.depends('line_ids.montant_a_facturer')
    def _compute_total_amount(self):
        """Calcule le montant total à facturer depuis les lignes"""
        for request in self:
            request.total_amount = sum(request.line_ids.mapped('montant_a_facturer'))

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if vals.get('name', 'Nouveau') == 'Nouveau':
                vals['name'] = self.env['ir.sequence'].next_by_code('project.invoice.request') or 'Nouveau'
        return super().create(vals_list)
    

    
    def action_submit(self):
        """Soumet la demande"""
        _logger.debug("=== DEBUG action_submit ===")
        _logger.debug(f"Utilisateur: {self.env.user.name} (ID: {self.env.user.id})")
        _logger.debug(f"Demandes à traiter: {self.mapped('name')}")

        for request in self:
            _logger.debug(f"État avant: {request.state}")
            if request.state == 'draft':
                request.write({'state': 'submitted'})
                _logger.debug(f"État après: {request.state}")
                request._send_email_submitted()
            else:
                _logger.debug(f"État incorrect pour soumission: {request.state}")
        return True

    def action_approve(self):
        """Approuve la demande"""
        _logger.debug("=== DEBUG action_approve ===")
        _logger.debug(f"Utilisateur: {self.env.user.name} (ID: {self.env.user.id})")

        # Vérifier les permissions
        can_validate = self.env.user.has_group('project_invoice_request.group_project_invoice_validator')
        _logger.debug(f"Utilisateur peut valider: {can_validate}")

        if not can_validate:
            raise UserError("Seuls les validateurs peuvent approuver les demandes.")

        _logger.debug(f"Demandes à traiter: {self.mapped('name')}")

        for request in self:
            _logger.debug(f"État avant: {request.state}")
            if request.state == 'submitted':
                request.write({'state': 'approved'})
                _logger.debug(f"État après: {request.state}")
                request._send_email_approved()
            else:
                _logger.debug(f"État incorrect pour approbation: {request.state}")
        return True

    def action_reject(self):
        """Rejette la demande"""
        _logger.debug("=== DEBUG action_reject ===")
        _logger.debug(f"Utilisateur: {self.env.user.name} (ID: {self.env.user.id})")

        # Vérifier les permissions
        can_validate = self.env.user.has_group('project_invoice_request.group_project_invoice_validator')
        _logger.debug(f"Utilisateur peut rejeter: {can_validate}")

        if not can_validate:
            raise UserError("Seuls les validateurs peuvent rejeter les demandes.")

        _logger.debug(f"Demandes à traiter: {self.mapped('name')}")

        for request in self:
            _logger.debug(f"État avant: {request.state}")
            if request.state == 'submitted':
                request.write({'state': 'rejected'})
                _logger.debug(f"État après: {request.state}")
                request._send_email_rejected()
            else:
                _logger.debug(f"État incorrect pour rejet: {request.state}")
        return True
    
    def action_reset_to_draft(self):
        """Remet la demande à l'état brouillon"""
        for request in self:
            if request.state not in ['submitted', 'rejected']:
                continue
            request.write({'state': 'draft'})
        return True

    def _send_email_submitted(self):
        """Envoi email aux validateurs avec documents"""
        _logger.info("=== DÉBUT _send_email_submitted ===")
        
        for request in self:
            # Récupérer les validateurs
            validator_group = self.env.ref('project_invoice_request.group_project_invoice_validator')
            validator_users = validator_group.users
            validator_emails = validator_users.mapped('email')
            validator_partner_ids = validator_users.mapped('partner_id').ids
            
            _logger.info("Validateurs: %s - Emails: %s", validator_users.mapped('name'), validator_emails)
            
            if not validator_emails:
                _logger.warning("Aucun validateur avec email configuré!")
                return
            
            # Préparer la section documents
            documents_html = ""
            if request.attachment_ids:
                documents_html = f"""
                <div style="background: #e9ecef; padding: 10px; border-radius: 5px; margin: 10px 0;">
                    <p><strong>📎 Documents joints ({len(request.attachment_ids)}) :</strong></p>
                    <ul>
                        {"".join([f'<li>{attachment.name}</li>' for attachment in request.attachment_ids])}
                    </ul>
                </div>
                """
            
            try:
                mail_values = {
                    'subject': f'Nouvelle Demande de Facturation - {request.name}',
                    'email_from': self.env.user.email or self.env.company.email,
                    'email_to': ','.join(validator_emails),
                    'body_html': f"""
                    <div style="font-family: Arial, sans-serif;">
                        <h3 style="color: #17a2b8;">📋 Nouvelle Demande de Facturation</h3>
                        <p>Une nouvelle demande de facturation a été soumise pour validation.</p>
                        
                        <div style="background: #f8f9fa; padding: 10px; border-radius: 5px;">
                            <p><strong>Détails :</strong></p>
                            <ul>
                                <li><strong>Référence :</strong> {request.name}</li>
                                <li><strong>Projet :</strong> {request.project_id.name}</li>
                                <li><strong>Commande Client :</strong> {request.sale_order_id.name}</li>                       
                                <li><strong>Client :</strong> {request.sale_order_id.partner_id.name}</li>
                                <li><strong>Montant :</strong> {request.total_amount} {request.currency_id.symbol}</li>
                                <li><strong>Soumise par :</strong> {request.create_uid.name}</li>
                            </ul>
                        </div>
                        
                        {documents_html}
                        
                        <p style="margin-top: 20px;">
                            <a href="{request.get_base_url()}/web#id={request.id}&model=project.invoice.request&view_type=form" 
                            style="background: #17a2b8; color: white; padding: 10px 15px; text-decoration: none; border-radius: 3px;">
                            Voir la demande
                            </a>
                        </p>
                    </div>
                    """,
                    'model': 'project.invoice.request',
                    'res_id': request.id,
                    'attachment_ids': [(6, 0, request.attachment_ids.ids)]  # Joindre les documents
                }
                
                # Créer et envoyer l'email
                mail = self.env['mail.mail'].create(mail_values)
                mail.send()
                _logger.info("Email créé et envoyé avec %s documents - ID: %s", len(request.attachment_ids), mail.id)
                
                # Ajouter au fil de discussion
                request.message_post(
                    body=f"""
                    <p><strong>Demande soumise pour validation</strong></p>
                    <p><strong>Documents joints :</strong> {len(request.attachment_ids)}</p>
                    <p>Email envoyé aux validateurs : {', '.join(validator_emails)}</p>
                    """,
                    subject="Demande soumise",
                    message_type='comment',
                    partner_ids=validator_partner_ids,
                    attachment_ids=request.attachment_ids.ids
                )
                _logger.info("Notification ajoutée au fil de discussion avec documents")
                
            except Exception as e:
                _logger.error("ERREUR envoi email: %s", str(e), exc_info=True)

    def _send_email_approved(self):
        """Envoi email aux comptables avec documents"""
        _logger.info("=== DÉBUT _send_email_approved ===")
        
        for request in self:
            # Récupérer les comptables
            accountant_group = self.env.ref('project_invoice_request.group_project_invoice_accountant')
            accountant_users = accountant_group.users
            accountant_emails = accountant_users.mapped('email')
            accountant_partner_ids = accountant_users.mapped('partner_id').ids
            
            _logger.info("Comptables: %s - Emails: %s", accountant_users.mapped('name'), accountant_emails)
            
            if not accountant_emails:
                _logger.warning("Aucun comptable avec email configuré!")
                return
            
            # Calcul du montant total
            montant_total = sum(request.line_ids.mapped('montant_a_facturer'))
            approver_name = self.env.user.name
            
            # Préparer la section documents
            documents_html = ""
            if request.attachment_ids:
                documents_html = f"""
                <div style="background: #e9ecef; padding: 10px; border-radius: 5px; margin: 10px 0;">
                    <p><strong>📎 Documents joints ({len(request.attachment_ids)}) :</strong></p>
                    <ul>
                        {"".join([f'<li>{attachment.name}</li>' for attachment in request.attachment_ids])}
                    </ul>
                </div>
                """
            
            try:
                mail_values = {
                    'subject': f'✅ Demande de Facturation Approuvée - {request.name}',
                    'email_from': self.env.user.email or self.env.company.email,
                    'email_to': ','.join(accountant_emails),
                    'body_html': f"""
                    <div style="font-family: Arial, sans-serif;">
                        <h3 style="color: #28a745;">✅ Demande Approuvée</h3>
                        <p>Cette demande de facturation a été approuvée par {approver_name} et nécessite un traitement comptable.</p>
                        
                        <div style="background: #f8f9fa; padding: 15px; border-radius: 5px;">
                            <p><strong>Détails :</strong></p>
                            <ul>
                                <li><strong>Référence :</strong> {request.name}</li>
                                <li><strong>Projet :</strong> {request.project_id.name}</li>
                                <li><strong>Commande Client :</strong> {request.sale_order_id.name}</li>
                                <li><strong>Client :</strong> {request.sale_order_id.partner_id.name}</li>
                                <li><strong>Montant à facturer :</strong> <span style="color: #28a745;">{montant_total} {request.currency_id.symbol}</span></li>
                                <li><strong>Approuvée par :</strong> {self.env.user.name}</li>
                            </ul>
                        </div>
                        
                        {documents_html}
                        
                        <p style="margin-top: 20px;">
                            <a href="{request.get_base_url()}/web#id={request.id}&model=project.invoice.request&view_type=form" 
                            style="background: #28a745; color: white; padding: 10px 15px; text-decoration: none; border-radius: 3px;">
                            Procéder à la facturation
                            </a>
                        </p>
                    </div>
                    """,
                    'model': 'project.invoice.request',
                    'res_id': request.id,
                    'attachment_ids': [(6, 0, request.attachment_ids.ids)]  # Joindre les documents
                }
                
                # Créer et envoyer l'email
                mail = self.env['mail.mail'].create(mail_values)
                mail.send()
                _logger.info("Email créé et envoyé avec %s documents - ID: %s", len(request.attachment_ids), mail.id)
                
                # Ajouter au fil de discussion
                request.message_post(
                    body=f"""
                    <p><strong>Demande approuvée</strong></p>
                    <p><strong>Montant à facturer :</strong> {montant_total} {request.currency_id.symbol}</p>
                    <p><strong>Documents joints :</strong> {len(request.attachment_ids)}</p>
                    <p>Email envoyé aux comptables : {', '.join(accountant_emails)}</p>
                    """,
                    subject="Demande approuvée",
                    message_type='comment',
                    partner_ids=accountant_partner_ids,
                    attachment_ids=request.attachment_ids.ids
                )
                _logger.info("Notification ajoutée au fil de discussion avec documents")
                
            except Exception as e:
                _logger.error("ERREUR envoi email: %s", str(e), exc_info=True)
                
    def _send_email_rejected(self):
        """Envoi email de rejet avec documents"""
        _logger.info("=== DÉBUT _send_email_rejected ===")
        
        for request in self:
            # Récupérer l'auteur de la demande
            author = request.create_uid
            author_email = author.email
            author_partner_ids = [author.partner_id.id] if author.partner_id else []
            
            _logger.info("Auteur de la demande: %s (Email: %s)", author.name, author_email)
            
            if not author_email:
                _logger.warning("L'auteur de la demande n'a pas d'email configuré!")
                request.message_post(
                    body=f"""
                    <p><strong>Demande rejetée</strong></p>
                    <p><strong>Documents joints :</strong> {len(request.attachment_ids)}</p>
                    <p>Rejetée par : {self.env.user.name}</p>
                    <p><em>Impossible d'envoyer l'email de notification</em></p>
                    """,
                    subject="Demande rejetée",
                    message_type='comment'
                )
                return
            
            # Préparer la section documents
            documents_html = ""
            if request.attachment_ids:
                documents_html = f"""
                <div style="background: #e9ecef; padding: 10px; border-radius: 5px; margin: 10px 0;">
                    <p><strong>📎 Documents joints ({len(request.attachment_ids)}) :</strong></p>
                    <ul>
                        {"".join([f'<li>{attachment.name}</li>' for attachment in request.attachment_ids])}
                    </ul>
                </div>
                """
            
            try:
                mail_values = {
                    'subject': f'❌ Demande de Facturation Rejetée - {request.name}',
                    'email_from': self.env.user.email or self.env.company.email,
                    'email_to': author_email,
                    'body_html': f"""
                    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                        <h2 style="color: #dc3545;">❌ Demande Rejetée</h2>
                        
                        <p>Bonjour <strong>{author.name}</strong>,</p>
                        
                        <p>Votre demande de facturation a été <strong>rejetée</strong> par le validateur.</p>
                        
                        <div style="background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 15px 0;">
                            <h3 style="margin-top: 0; color: #495057;">Détails de la demande :</h3>
                            <table style="width: 100%;">
                                <tr>
                                    <td style="padding: 5px; font-weight: bold; width: 120px;">Référence :</td>
                                    <td style="padding: 5px;"><strong>{request.name}</strong></td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px; font-weight: bold;">Projet :</td>
                                    <td style="padding: 5px;">{request.project_id.name}</td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px; font-weight: bold;">Client :</td>
                                    <td style="padding: 5px;">{request.sale_order_id.partner_id.name}</td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px; font-weight: bold;">Montant :</td>
                                    <td style="padding: 5px;"><strong>{request.total_amount} {request.currency_id.symbol}</strong></td>
                                </tr>
                                <tr>
                                    <td style="padding: 5px; font-weight: bold;">Rejetée par :</td>
                                    <td style="padding: 5px;">{self.env.user.name}</td>
                                </tr>
                            </table>
                        </div>
                        
                        {documents_html}
                        
                        <div style="background: #f8d7da; padding: 15px; border-radius: 5px; margin: 15px 0;">
                            <h4 style="margin-top: 0; color: #721c24;">📞 Contactez le validateur :</h4>
                            <p>
                                Pour plus d'informations, contactez :<br/>
                                <strong>{self.env.user.name}</strong> - <a href="mailto:{self.env.user.email}">{self.env.user.email}</a>
                            </p>
                        </div>
                        
                        <div style="text-align: center; margin: 20px 0;">
                            <a href="{request.get_base_url()}/web#id={request.id}&model=project.invoice.request&view_type=form" 
                            style="background: #6c757d; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">
                            📋 Voir la demande
                            </a>
                        </div>
                    </div>
                    """,
                    'model': 'project.invoice.request',
                    'res_id': request.id,
                    'attachment_ids': [(6, 0, request.attachment_ids.ids)]  # Joindre les documents
                }
                
                # Créer et envoyer l'email
                mail = self.env['mail.mail'].create(mail_values)
                mail.send()
                _logger.info("Email de rejet créé et envoyé avec %s documents - ID: %s", len(request.attachment_ids), mail.id)
                
                # Ajouter au fil de discussion
                request.message_post(
                    body=f"""
                    <div style="font-family: Arial, sans-serif;">
                        <p><strong>❌ Demande rejetée</strong></p>
                        <p><strong>Documents joints :</strong> {len(request.attachment_ids)}</p>
                        <p>Rejetée par : <strong>{self.env.user.name}</strong></p>
                        <p>Email envoyé à : <strong>{author.name}</strong></p>
                    </div>
                    """,
                    subject="Demande rejetée",
                    message_type='comment',
                    partner_ids=author_partner_ids,
                    attachment_ids=request.attachment_ids.ids
                )
                _logger.info("Notification de rejet ajoutée au fil de discussion avec documents")
                
            except Exception as e:
                _logger.error("ERREUR envoi email de rejet: %s", str(e), exc_info=True)
                request.message_post(
                    body=f"""
                    <p><strong>Demande rejetée</strong></p>
                    <p><strong>Documents joints :</strong> {len(request.attachment_ids)}</p>
                    <p>Rejetée par : {self.env.user.name}</p>
                    <p style="color: #dc3545;"><em>Erreur lors de l'envoi de l'email</em></p>
                    """,
                    subject="Demande rejetée",
                    message_type='comment'
                )
                        
    def action_create_invoice(self):
        """Crée la facture à partir de la demande approuvée"""
        self.ensure_one()
        
        if self.state != 'approved':
            raise UserError("Seules les demandes approuvées peuvent être facturées.")
        
        if self.invoice_id:
            raise UserError("Une facture a déjà été créée pour cette demande.")
        
        # CRÉATION SIMPLE de facture avec montant total
        invoice_vals = {
            'move_type': 'out_invoice',
            'partner_id': self.sale_order_id.partner_id.id,
            'invoice_origin': f"{self.sale_order_id.name} - {self.name}",
            'invoice_line_ids': [(0, 0, {
                'name': f"Facturation projet {self.project_id.name} - {self.description or 'Sans description'}",
                'quantity': 1,
                'price_unit': self.total_amount,  # Utiliser le montant total
            })],
        }
        
        invoice = self.env['account.move'].create(invoice_vals)
        self.write({
            'invoice_id': invoice.id,
            'state': 'invoiced',
        })
        
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'account.move',
            'res_id': invoice.id,
            'view_mode': 'form',
            'target': 'current',
        }
    
    def test_email_configuration(self):
        """Méthode de test pour la configuration email"""
        _logger.info("=== TEST CONFIGURATION EMAIL ===")
        
        # Vérifier les serveurs SMTP
        smtp_servers = self.env['ir.mail_server'].search([])
        _logger.info("Servers SMTP configurés: %s", len(smtp_servers))
        for server in smtp_servers:
            _logger.info("Server SMTP: %s - %s:%s", server.name, server.smtp_host, server.smtp_port)
        
        # Vérifier les groupes
        validators = self.env.ref('project_invoice_request.group_project_invoice_validator', raise_if_not_found=False)
        _logger.info("Groupe validateurs trouvé: %s", bool(validators))
        if validators:
            _logger.info("Validateurs: %s", validators.users.mapped('name'))
            _logger.info("Emails validateurs: %s", validators.users.mapped('email'))
        
        accountants = self.env.ref('project_invoice_request.group_project_invoice_accountant', raise_if_not_found=False)
        _logger.info("Groupe comptables trouvé: %s", bool(accountants))
        if accountants:
            _logger.info("Comptables: %s", accountants.users.mapped('name'))
            _logger.info("Emails comptables: %s", accountants.users.mapped('email'))
        
        # Tester un envoi d'email simple
        try:
            test_email = self.env['mail.mail'].create({
                'subject': 'Test configuration email',
                'email_to': self.env.user.email,
                'body_html': '<p>Ceci est un test de configuration email.</p>',
            })
            test_email.send()
            _logger.info("Email de test envoyé avec succès à: %s", self.env.user.email)
        except Exception as e:
            _logger.error("ERREUR envoi email test: %s", str(e), exc_info=True)
        
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': 'Test Configuration Email',
                'message': 'Vérifiez les logs pour les résultats du test.',
                'type': 'info',
                'sticky': False,
            }
        }


class ProjectInvoiceRequestLine(models.Model):
    _name = 'project.invoice.request.line'
    _description = 'Ligne de Demande de Facturation'
    
    request_id = fields.Many2one('project.invoice.request', required=True, ondelete='cascade')
    
    # GARDER uniquement les champs montants
    montant_facture = fields.Monetary(string='Montant Déjà Facturé', currency_field='currency_id')
    montant_restant = fields.Monetary(string='Montant Restant', currency_field='currency_id')
    montant_a_facturer = fields.Monetary(string='Montant à Facturer', required=True, currency_field='currency_id')
    
    currency_id = fields.Many2one(related='request_id.currency_id')