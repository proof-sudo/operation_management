from odoo import models, fields, api

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'
    dossier_id = fields.Char(string='Dossier', copy=False)
    # date_previsionnelle_livraison = fields.Datetime(string='Date Prévisionnelle de Livraison')
    eta_constructeur = fields.Datetime(string='ETA Constructeur ')
    instructions_speciales = fields.Text(string='Instructions Spéciales')
    statut_livraison = fields.Selection([
        ('en_attente', 'En Attente'),
        ('partiellement_livre', 'Partiellement Livré'),
        ('livre', 'Livré'),
        ('annule', 'Annulé'),
        ('placee', 'Placée')
    ], string='Statut de Livraison', default='en_attente')
    date_planned = fields.Datetime(string='Date Prévisionnelle de Livraison')
    ref_Fp = fields.Char(string='Commande Client / FP')
    client_id = fields.Many2one(string='Client', comodel_name='res.partner', ondelete='restrict')
    
# class DossierCommercial(models.Model):
#     _inherit = 'dossier.commercial'
    
#     purchase_order_ids = fields.One2many('purchase.order', 'dossier_id', string='Commandes d\'Achat Associées')
#     client_id = fields.Many2one(string='Client', comodel_name='res.partner', ondelete='restrict')
#     ref_bon_commande_client = fields.Char(string='Référence Bon de Commande Client')
    

  