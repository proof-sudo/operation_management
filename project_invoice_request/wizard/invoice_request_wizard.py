from odoo import models, fields, api
from odoo.exceptions import UserError, ValidationError

class ProjectInvoiceRequestWizard(models.TransientModel):
    _name = 'project.invoice.request.wizard'
    _description = 'Wizard de Demande de Facturation'
    
    project_id = fields.Many2one('project.project', string='Projet', required=True)
    sale_order_id = fields.Many2one('sale.order', string='Commande Client', required=True)
    
    # Champs pour les montants
    montant_total_bc = fields.Monetary(
        string='Montant Total BC', 
        related='sale_order_id.amount_total',
        currency_field='currency_id',
        readonly=True
    )
    montant_deja_facture = fields.Monetary(
        string='Montant Déjà Facturé',
        compute='_compute_montants',
        currency_field='currency_id',
        readonly=True
    )
    montant_disponible = fields.Monetary(
        string='Montant Restant',
        compute='_compute_montants', 
        currency_field='currency_id',
        readonly=True
    )
    montant_a_facturer = fields.Monetary(
        string='Montant à Facturer',
        currency_field='currency_id',
        required=True
    )
    currency_id = fields.Many2one(
        'res.currency',
        related='sale_order_id.currency_id',
        readonly=True
    )
    
    # Champ pour description/notes
    description = fields.Text(string='Description / Notes')
    
    # Champ pour documents
    attachment_ids = fields.Many2many(
        'ir.attachment',
        'project_invoice_request_wizard_attachment_rel',
        'wizard_id', 
        'attachment_id',
        string='Documents à Joindre'
    )

    @api.depends('sale_order_id')
    def _compute_montants(self):
        """Calcule les montants déjà facturés et disponibles"""
        for wizard in self:
            if wizard.sale_order_id:
                # Récupérer toutes les demandes de facturation approuvées/invoiced pour cette commande
                existing_requests = self.env['project.invoice.request'].search([
                    ('sale_order_id', '=', wizard.sale_order_id.id),
                    ('state', 'in', ['approved', 'invoiced']),
                ])
                
                # Calcul du montant déjà facturé
                wizard.montant_deja_facture = sum(existing_requests.mapped('total_amount'))
                
                # Calcul du montant disponible
                wizard.montant_disponible = wizard.montant_total_bc - wizard.montant_deja_facture
            else:
                wizard.montant_deja_facture = 0.0
                wizard.montant_disponible = 0.0

    @api.constrains('montant_a_facturer')
    def _check_montant_a_facturer(self):
        """Vérifie que le montant à facturer est valide"""
        for wizard in self:
            if wizard.montant_a_facturer <= 0:
                raise ValidationError("Le montant à facturer doit être supérieur à 0.")
            if wizard.montant_a_facturer > wizard.montant_disponible:
                raise ValidationError(
                    f"Le montant à facturer ({wizard.montant_a_facturer}) "
                    f"dépasse le montant disponible ({wizard.montant_disponible})."
                )

    def action_submit_request(self):
        """Soumet la demande de facturation basée sur les montants"""
        self.ensure_one()
        
        # Vérification du montant
        if self.montant_a_facturer <= 0:
            raise UserError("Veuillez saisir un montant à facturer supérieur à 0.")
        
        if self.montant_a_facturer > self.montant_disponible:
            raise UserError(
                f"Le montant à facturer ({self.montant_a_facturer}) "
                f"dépasse le montant disponible ({self.montant_disponible})."
            )
        
        # CRÉATION SIMPLIFIÉE - une seule ligne cohérente
        request = self.env['project.invoice.request'].create({
            'project_id': self.project_id.id,
            'sale_order_id': self.sale_order_id.id,
            'state': 'draft',
            'description': self.description,
            'attachment_ids': [(6, 0, self.attachment_ids.ids)],
            'line_ids': [(0, 0, {
                'montant_facture': self.montant_deja_facture,
                'montant_restant': self.montant_disponible,
                'montant_a_facturer': self.montant_a_facturer,
            })]
        })
        
        # Soumettre automatiquement la demande
        request.action_submit()
        
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'project.invoice.request',
            'res_id': request.id,
            'view_mode': 'form',
            'target': 'current',
        }
    # def _get_default_product(self):
    #     """Retourne un produit par défaut pour les facturations"""
    #     product = self.env['product.product'].search([
    #         ('default_code', '=', 'FACTURATION_PROJET'),
    #         ('type', '=', 'service')
    #     ], limit=1)
        
    #     if not product:
    #         # Créer un produit par défaut si nécessaire
    #         product = self.env['product.product'].create({
    #             'name': 'Facturation Projet',
    #             'default_code': 'FACTURATION_PROJET',
    #             'type': 'service',
    #             'invoice_policy': 'order',
    #             'sale_ok': True,
    #             'purchase_ok': False,
    #             'list_price': 0.0,
    #         })
        
    #     return product

# SUPPRIMER la classe ProjectInvoiceRequestWizardLine - PLUS UTILE