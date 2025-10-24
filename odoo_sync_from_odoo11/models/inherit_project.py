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
    am = fields.Many2one('res.users', string='Account Manager', related='bc.user_id',store=True,readonly=False)
    presales = fields.Many2one('res.users', string='Presales',store=True,readonly=False)
    date_in = fields.Date(string='Date IN', compute='_compute_creation_date_only',store=True,readonly=False )
    pays = fields.Many2one('res.country', string='Pays', related='bc.partner_id.country_id',store=True,readonly=False )
   
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
    etat_projet = fields.Selection([
        ('cancelled', '0-Annulé'),
        ('dossier_indisponible', '6-Dossier indisponible'),
        ('suspendu', '9-Suspendu'),
        ('draft', '6-Draft'),
        ('non_demarre', '1-Non démarré'),
        ('en_cours_bloque', '3-En cours - Bloqué'),
        ('en_cours_provisionning', '3-En cours - Provisionning'),
        ('en_cours_production', '3-En cours - Production'),
        ('en_cours_expedition', '3-En cours - Expédition'),
        ('en_cours_dedouanement', '3-En cours - Dedouanement'),
        ('en_cours_atelier_technique', '3-En cours - Atelier technique'),
        ('en_cours_deploiement', '3-En cours - Deploiement'),
        ('en_cours_formation', '3-En cours - Formation'),
        ('en_cours_kickoff_client', '3-En cours - Kick off client'),
        ('en_cours_standby_client', '3-En cours - Standby client'),
        ('en_cours_standby_technical_issue', '3-En cours - Standby technical issue'),
        ('en_cours_attente_prerequis', '3-En cours - Attente prérequis'),
        ('en_cours_tests_recette', '3-En cours - Tests et recette'),
        ('en_cours_rli', '3-En cours - RLI'),
        ('termine_attente_pv_bl', '4-Terminé - Attente PV/BL'),
        ('termine_levee_reserve', '4-Terminé - Lévée de reserve'),
        ('termine_pv_bl_signe', '4-Terminé - PV/BL signé'),
        ('facture_attente_df', '5-Facturé - Attente DF'),
        ('facture_attente_livraison', '5-Facturé - Attente livraison'),
        ('facture_prestations_en_cours', '5-Facturé - Prestations en cours'),
        ('suivi_contrat_licence', '8-Suivi - Contrat licence'),
        ('suivi_contrat_mixte', '8-Suivi - Contrat Mixte'),
        ('suivi_contrat_services', '8-Suivi - Contrat de Services'),
        ('cloture', '7-Cloturé'),
    ], string='État projet', default='non_demarre', help='État détaillé du projet')
    
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
    cas_hw =fields.Float(string='CAS HW', default=0.0)
    secteur = fields.Many2one(
        'res.partner.category',
        string='Secteur',
        compute='_compute_secteur_from_bc',
        inverse='_inverse_secteur_to_partner',
        store=True
    )
    
    @api.depends('bc.partner_id.category_id', 'partner_id.category_id')
    def _compute_secteur_from_bc(self):
        for record in self:
            # Priorité 1 : Commande liée (BC)
            if record.bc and record.bc.partner_id.category_id:
                record.secteur = record.bc.partner_id.category_id[0]
            # Priorité 2 : Partenaire direct du projet
            elif record.partner_id and record.partner_id.category_id:
                record.secteur = record.partner_id.category_id[0]
            else:
                record.secteur = False
    
    def _inverse_secteur_to_partner(self):
        """Quand on modifie le secteur depuis le projet, on met à jour le partenaire"""
        for record in self:
            if record.partner_id and record.secteur:
                record.partner_id.category_id = [(6, 0, [record.secteur.id])]
    # secteur= fields.Char(string='Secteur', compute='_compute_secteur', store=True)
    
    
    
    # @api.depends('bc.partner_id')
    # def _compute_secteur(self):
    #     for project in self:
    #         if project.bc and project.bc.partner_id:
    #             project.secteur = project.bc.partner_id.secteur
    #         else:
    #             project.secteur = ''

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
