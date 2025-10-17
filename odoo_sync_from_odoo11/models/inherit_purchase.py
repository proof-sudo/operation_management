from odoo import models, fields, api

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'
    dossier_id = fields.Char(string='Dossier', copy=False)
    # date_previsionnelle_livraison = fields.Datetime(string='Date Prévisionnelle de Livraison')
    date_enlevement = fields.Datetime(string='Date d\'Enlèvement ')
    instructions_speciales = fields.Text(string='Instructions Spéciales')
    statut_livraison = fields.Selection([
        ('en_attente', 'En Attente'),
        ('partiellement_livre', 'Partiellement Livré'),
        ('livre', 'Livré'),
        ('annule', 'Annulé'),
        ('placee', 'Placée')
    ], string='Statut de Livraison', default='en_attente')
    

  