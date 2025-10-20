from odoo import models, fields, api, _

class ProjectInherit(models.Model):
    _inherit = 'project.project'
    
    nature = fields.Selection([
        ('all', 'ALL'),
        ('end_to_end', 'End to End'),
        ('livraison', 'Livraison'),
        ('service_pro', 'Service Pro'),
    ], string='Nature', default='all')

    domaine = fields.Selection([
        ('others', 'Others'),
        ('datacenter_facilities', 'Datacenter Facilities (DCF)'),
        ('modern_network_integration', 'Modern Network Integration (MNI)'),
        ('agile_infrastructure_cloud', 'Agile Infrastructure & Cloud (AIC)'),
        ('business_data_integration', 'Business Data Integration (BDI)'),
        ('digital_workspace', 'Digital Workspace (DWS)'),
        ('secured_it', 'Secured IT (SEC)'),
        ('expert_managed_services_think', 'Expert & Managed Services - THINK'),
        ('expert_managed_services_build', 'Expert & Managed Services - BUILD'),
        ('expert_managed_services_train', 'Expert & Managed Services - TRAIN'),
        ('expert_managed_services_run', 'Expert & Managed Services - RUN'),
        ('none', 'None')
    ], string='Domaine', default='others')
    bc = fields.Many2one('sale.order', string='Commande liée', help="Commande liée à ce projet")
    am = fields.Many2one('res.users', string='AM', related='bc.user_id', store=True, readonly=True)
    presales = fields.Many2one('res.users', string='Presales')
    date_in = fields.Date(string='Date IN', compute='_compute_creation_date_only', store=True, readonly=True)
    pays = fields.Many2one('res.country', string='Pays', related='bc.partner_id.country_id', store=True, readonly=True)
    circuit = fields.Selection(string='Circuit', selection=[('fast', 'Fast Track'), ('normal', 'Normal')], default='normal')
    sc = fields.Many2one('res.users', string='Solutions consultant')
    cas = fields.Float(string='CAS', default=0.0)
    revenue_type = fields.Selection([
        ('oneshot', 'One Shot'),
        ('recurrent', 'Recurrent'),
    ], string='Revenue', default='oneshot')
    cafy = fields.Float(string='CAF YTD', default=0.0)
    rafytd = fields.Float(string='Raf YTD', default=0.0)
    cafypercent = fields.Float(string='CAF YTD %')
    rafy_1=fields.Float(string='Raf Y+1', default=0.0)
    projected_caf_y = fields.Float(string='Projected CAF Y', default=0.0)
    raftotal = fields.Float(string='Raf Total', default=0.0)
    percentcaftotal = fields.Float(string='% CAF Total')
    risque = fields.Selection([('delay', 'Delai'), ('cost', 'Cout'), ('quality', 'Qualité'), ('scope', 'Périmètre')], string='Risque')
    last_notice_date = fields.Date(string='Last Notice Date')
    contratstartdate = fields.Date(string='Contrat Start Date')
    contratenddate = fields.Date(string='Contrat End Date')
    delaicontractuel = fields.Date(string='Délai Contractuel')
    priorite = fields.Selection([('urgent', 'Urgent'), ('normal', 'Normal'), ('basse', 'Basse')], string='Priorité', default='normal')

    bu  = fields.Selection([('ict', 'ICT'), 
                            ('cloud', 'CLOUD'),
                            ('cybersecurity', 'CYBERSECURITY'),
                            ('formation', 'FORMATION'),
                            ('security', 'SECURITY')], string='BU')
    cat_recurrent = fields.Char(string='Cat Recurrent')
    cas_build =fields.Float(string='CAS BUILD', default=0.0)
    cas_run =fields.Float(string='CAS RUN', default=0.0)
    cas_train =fields.Float(string='CAS TRAIN', default=0.0)
    cas_sw =fields.Float(string='CAS SW', default=0.0)


    @api.depends('bc')
    def _compute_cas(self):
        for project in self:
            if project.bc:
                total_cas =  project.bc.amount_total
                project.cas = total_cas
            else:
                project.cas = 0.0
    @api.depends('create_date')
    def _compute_creation_date_only(self):
        for project in self:
            if project.create_date:
                project.date_in = project.create_date.date()
            else:
                project.date_in = False

    # @api.model
    # def create(self, vals):
    #     if 'sale_order_id' in vals and vals['sale_order_id']:
    #         sale_order = self.env['sale.order'].browse(vals['sale_order_id'])
    #         vals['name'] = f"Projet pour {sale_order.name}"
    #     return super(Project, self).create(vals)
