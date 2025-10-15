# sale_project_creation/wizards/create_project_wizard.py
from odoo import models, fields, api

class CreateProjectWizard(models.TransientModel):
    _name = 'create.project.wizard'
    _description = 'Assistant de création de projet depuis une vente'

    # Champ pour le nom du projet, pré-rempli avec le nom de la vente
    name = fields.Char(string="Nom du Projet", related='sale_order_id.project_name')
    
    # Champs pour stocker les informations de la vente d'origine
    sale_order_id = fields.Many2one('sale.order', string="Commande d'origine", readonly=True)
    sale_order_reference = fields.Char(string="Référence de la Commande", readonly=True)
    amount_total = fields.Monetary(string="Montant de la Commande", readonly=True)
    currency_id = fields.Many2one('res.currency', related='sale_order_id.currency_id', readonly=True)
    chef_de_projet = fields.Many2one('res.users', string="Chef de Projet", default=lambda self: self.env.user)
    circuit = fields.Selection(string='Circuit', selection=[('fast', 'Fast Track'), ('normal', 'Normal')], default='normal')
    delaicontractuel = fields.Date(string='Délai Contractuel')
    priorite = fields.Selection([('urgent', 'Urgent'), ('normal', 'Normal'), ('basse', 'Basse')], string='Priorité', default='normal')

    # Méthode appelée par le bouton "Créer le Projet" du wizard
    def action_create_project(self):
        self.ensure_one()
        
        # Logique de création du projet
        project = self.env['project.project'].create({
            'name': self.sale_order_id.project_name or 'Projet pour ' + self.sale_order_reference,
            'sale_order_id': self.sale_order_id.id,
            'user_id': self.chef_de_projet.id,
            'partner_id': self.sale_order_id.partner_id.id,
            'reinvoiced_sale_order_id': self.sale_order_id.id,
            'bc': self.sale_order_id.id,
            'circuit': self.circuit,
            'delaicontractuel': self.delaicontractuel,
            'priorite': self.priorite,
            # Vous pouvez mapper d'autres champs ici. Par exemple :
            # 'partner_id': self.sale_order_id.partner_id.id,
        })
        
        # Retourne une action pour ouvrir le projet nouvellement créé
        return {
            'type': 'ir.actions.act_window',
            'name': 'Projet créé',
            'res_model': 'project.project',
            'res_id': project.id,
            'view_mode': 'form',
            'target': 'current', # Ouvre dans la fenêtre principale, pas un nouveau pop-up
        }