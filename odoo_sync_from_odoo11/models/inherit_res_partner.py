from odoo import models, fields, api, _

class ResPartnerInherit(models.Model):
    _inherit = 'res.partner'
    
    secteur = fields.Char(string='Secteur')
    