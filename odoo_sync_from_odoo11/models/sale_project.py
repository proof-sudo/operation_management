from odoo import models, fields, api

class SaleOrder(models.Model):
    _inherit = 'sale.order'
    
    project_name = fields.Char(string="Nom du Projet", help="Nom du projet associé à cette commande")
    circuit = fields.Selection(string='Circuit', selection=[('fast', 'Fast Track'), ('normal', 'Normal')], default='normal')
    delaicontractuel = fields.Date(string='Délai Contractuel')
    priorite = fields.Selection([('urgent', 'Urgent'), ('normal', 'Normal'), ('basse', 'Basse')], string='Priorité', default='normal')

    def action_open_create_project_wizard(self):
        self.ensure_one()
        
        # Création de l'enregistrement du wizard avec les valeurs par défaut
        wizard = self.env['create.project.wizard'].create({
            'sale_order_id': self.id,
            'sale_order_reference': self.name,
            'amount_total': self.amount_total,
            'name': f"Projet pour {self.name}",
            'circuit': self.circuit,
            'delaicontractuel': self.delaicontractuel,
            'priorite': self.priorite,
            # Nom par défaut pour le projet
        })
        
        # Retourne l'action pour ouvrir la vue du wizard
        return {
            'type': 'ir.actions.act_window',
            'name': 'Créer un Projet depuis une Vente',
            'res_model': 'create.project.wizard',
            'view_mode': 'form',
            'res_id': wizard.id,
            'target': 'new', # 'new' pour ouvrir en pop-up
        }