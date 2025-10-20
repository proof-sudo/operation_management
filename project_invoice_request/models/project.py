from odoo import models, fields, api
from odoo.exceptions import UserError

class Project(models.Model):
    _inherit = 'project.project'
    
    invoice_request_ids = fields.One2many(
        'project.invoice.request', 
        'project_id', 
        string='Demandes de Facturation'
    )
    
    invoice_request_count = fields.Integer(
        string='Nombre de Demandes', 
        compute='_compute_invoice_request_count'
    )
    
    demande_facturation_count = fields.Integer(
        string='Nombre de Demandes', 
        compute='_compute_demande_facturation_count'
    )
    
    total_invoiced_amount = fields.Monetary(
        string='Montant Total Facturé',
        compute='_compute_total_invoiced_amount',
        currency_field='currency_id',
        help="Montant total déjà facturé pour ce projet (basé sur les demandes approuvées)"
    )
    
    total_backlog = fields.Monetary(
        string='Montant Total Backlog',
        compute='_compute_total_backlog',
        currency_field='currency_id',
        help="Montant backlog pour ce projet (basé sur les commandes liées)"
    )
    total_submit = fields.Monetary(
        string="En attente de validation",
        compute='_compute_total_submitted',
        currency_field='currency_id',
        help="Montant soumis en attente de valodation"
    )
    
    currency_id = fields.Many2one(
        'res.currency',
        string='Devise',
        compute='_compute_currency_id'
    )
    
    
    @api.depends('total_invoiced_amount')
    def _compute_total_backlog(self):
        """Calcule le montant total backlog basé sur les commandes liées"""
        for project in self:
            if project.bc:
                project.total_backlog = project.bc.amount_total - project.total_invoiced_amount
            else:
                project.total_backlog = 0.0
    
    def _compute_currency_id(self):
        """Calcule la devise basée sur la commande liée"""
        for project in self:
            if project.bc and project.bc.currency_id:
                project.currency_id = project.bc.currency_id
            else:
                project.currency_id = False
    
    def _compute_total_invoiced_amount(self):
        """Calcule le montant total facturé basé sur project.invoice.request"""
        for project in self:
            # Récupérer toutes les demandes approuvées pour ce projet
            approved_requests = self.env['project.invoice.request'].search([
                ('project_id', '=', project.id),
                ('state', '=', 'approved')
            ])
            
            # Calcul simple : somme des montants totaux des demandes
            total = sum(approved_requests.mapped('total_amount'))
            project.total_invoiced_amount = total
            
    def _compute_total_submitted(self):
        """Calcule le montant total facturé basé sur project.invoice.request"""
        for project in self:
            # Récupérer toutes les demandes approuvées pour ce projet
            approved_requests = self.env['project.invoice.request'].search([
                ('project_id', '=', project.id),
                ('state', '=', 'submitted')
            ])
            
            # Calcul simple : somme des montants totaux des demandes
            total = sum(approved_requests.mapped('total_amount'))
            project.total_submit = total

    @api.depends('invoice_request_ids')
    def _compute_invoice_request_count(self):
        for project in self:
            project.invoice_request_count = len(project.invoice_request_ids)

    @api.depends('invoice_request_ids')
    def _compute_demande_facturation_count(self):
        for project in self:
            project.demande_facturation_count = len(project.invoice_request_ids)

    def action_request_invoice(self):
        """Ouvre le wizard de demande de facturation"""
        self.ensure_one()
        
        if not self.bc:
            raise UserError("Aucune commande client associée à ce projet.")
        
        return {
            'name': 'Demande de Facturation',
            'type': 'ir.actions.act_window',
            'res_model': 'project.invoice.request.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_project_id': self.id,
                'default_sale_order_id': self.bc.id,
            }
        }
    
    def action_view_invoice_requests(self):
        """Affiche les demandes de facturation du projet"""
        self.ensure_one()
        return {
            'name': 'Demandes de Facturation',
            'type': 'ir.actions.act_window',
            'res_model': 'project.invoice.request',
            'view_mode': 'tree,form',
            'domain': [('project_id', '=', self.id)],
            'context': {'default_project_id': self.id}
        }