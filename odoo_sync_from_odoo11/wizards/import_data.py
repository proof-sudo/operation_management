from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
import io
import logging
from datetime import datetime
import re 

_logger = logging.getLogger(__name__)

try:
    import openpyxl
except ImportError:
    _logger.warning("Le module openpyxl n'est pas installé. Installation requise: pip install openpyxl")
    openpyxl = None

# Mappage des colonnes du fichier Excel vers les champs Odoo
COLUMN_MAPPING = {
    'Nom': 'name',
    'PM': 'user_id', # Project Manager (res.users)
    'AM': 'account_manager_id', 
    'Presales': 'presales_id', 
    'Nature': 'nature', # Sélection
    'BU': 'bu', # Sélection
    'Domaine': 'domaine', # Sélection
    'Revenus': 'revenue_type', 
    'Cat Recurrent': 'recurrent_category', 
    'Date IN': 'date_start',
    'Pays': 'country_id', 
    'Customer': 'partner_id', 
    'Secteur': 'partner_category_ids', # Renommé pour correspondre au Many2many standard
    'Description du Projet': 'description',
    'Circuit': 'circuit', # Sélection
    'SC': 'cost_sc', 
    'CAS Build': 'cost_cas_build',
    'CAS Run': 'cost_cas_run',
    'CAS Train': 'cost_cas_train',
    'CAS Sw': 'cost_cas_sw',
    'CAS Hw': 'cost_cas_hw',
    'CAS': 'cost_cas_total',
    'Statut': 'etat_projet', 
    'Update Date': 'date_update',
}


class ProjectImportWizard(models.TransientModel):
    _name = 'project.import.wizard'
    _description = "Wizard d'import de projets depuis Excel"

    import_file = fields.Binary(
        string='Fichier Excel',
        required=True
    )
    import_filename = fields.Char(string='Nom du fichier')
    
    update_existing = fields.Boolean(
        string='Mettre à jour les projets existants',
        default=True
    )
    create_missing = fields.Boolean(
        string='Créer les projets manquants',
        default=True
    )
    create_missing_records = fields.Boolean(
        string='Créer les enregistrements manquants (utilisateurs, clients, etc.)',
        default=True,
        help="Si activé, crée automatiquement les utilisateurs, clients et autres enregistrements manquants"
    )
    
    import_log = fields.Text(string="Journal d'import", readonly=True)
    success_count = fields.Integer(string='Projets créés/mis à jour', readonly=True)
    error_count = fields.Integer(string='Erreurs', readonly=True)
    
    # CHAMPS AJOUTÉS POUR CORRESPONDRE À LA VUE XML DE L'UTILISATEUR
    created_users_count = fields.Integer(string='Utilisateurs créés', readonly=True)
    created_partners_count = fields.Integer(string='Clients créés', readonly=True)
    created_categories_count = fields.Integer(string='Catégories créées', readonly=True)
    # FIN DES AJOUTS

    # --- MÉTHODES DE GESTION DES ENREGISTREMENTS EXTERNES ---
    
    def _find_or_create_user(self, name):
        """ Recherche ou crée un utilisateur, en assurant l'unicité du login et en évitant 'default'. """
        if not self.create_missing_records:
            return False
            
        name = str(name or '').strip()
        # Ignorer les noms vides ou génériques qui causent l'erreur 'default'
        if not name or name.lower() in ('nan', 'none', 'default', 'n/a', 'na'):
            return False

        User = self.env['res.users'].sudo()
        
        # 1. Tenter de trouver l'utilisateur par nom
        user = User.search(['|', ('name', '=ilike', name), ('login', '=ilike', name)], limit=1)
        if user:
            return user.id

        # 2. Créer l'utilisateur (avec gestion de login unique)
        try:
            # Nettoyage et création du login de base (ex: Fatim KEITA -> fatim.keita)
            login_base = re.sub(r'[^a-zA-Z0-9\.]', '', name.lower().replace(' ', '.'))
            if not login_base:
                 login_base = 'imported.user' 
                
            login_candidate = login_base
            login_suffix = 0
            
            # Garantir l'unicité du login
            while User.search([('login', '=', login_candidate)], limit=1):
                login_suffix += 1
                login_candidate = f"{login_base}{login_suffix}"
            
            user_vals = {
                'name': name,
                'login': login_candidate,
                'email': f'{login_candidate}@{self.env.company.partner_id.email_domain or "example.com"}',
                'company_id': self.env.company.id,
                'notification_type': 'email',
                'groups_id': [(6, 0, [self.env.ref('base.group_user').id])]
            }
            
            new_user = User.create(user_vals)
            self.created_users_count += 1 # Incrémentation du compteur
            return new_user.id
            
        except Exception as e:
            _logger.error("Erreur recherche/création res.users '%s': %s", name, str(e))
            self.import_log += _("Erreur: Impossible de créer l'utilisateur '%s': %s\n" % (name, str(e)))
            return False

    def _find_or_create_partner(self, name):
        """ Recherche ou crée un partenaire (Client). """
        if not self.create_missing_records:
            return False
            
        name = str(name or '').strip()
        if not name or name.lower() in ('nan', 'none', 'n/a', 'na'):
            return False
            
        Partner = self.env['res.partner'].sudo()
        partner = Partner.search([('name', '=ilike', name), ('is_company', '=', True)], limit=1)
        
        if partner:
            return partner.id
            
        # Création du partenaire
        try:
            new_partner = Partner.create({
                'name': name,
                'is_company': True,
                'company_type': 'company',
            })
            self.created_partners_count += 1 # Incrémentation du compteur
            return new_partner.id
        except Exception as e:
            _logger.error("Erreur recherche/création res.partner '%s': %s", name, str(e))
            self.import_log += _("Erreur: Impossible de créer le partenaire '%s': %s\n" % (name, str(e)))
            return False

    def _find_or_create_misc(self, model_name, name):
        """ Recherche ou crée d'autres enregistrements (Pays, Catégories, etc.). """
        if not self.create_missing_records:
            return False
            
        name = str(name or '').strip()
        if not name or name.lower() in ('nan', 'none', 'n/a', 'na'):
            return False

        Model = self.env[model_name].sudo()
        domain = [('name', '=ilike', name)]
        
        if model_name == 'res.country':
            domain = ['|', ('name', '=ilike', name), ('code', '=ilike', name)]
            
        record = Model.search(domain, limit=1)
        if record:
            return record.id

        # Création de l'enregistrement
        try:
            new_record = Model.create({'name': name})
            if model_name == 'res.partner.category':
                self.created_categories_count += 1 # Incrémentation du compteur
            return new_record.id
        except Exception as e:
            _logger.error("Erreur recherche/création %s '%s': %s", model_name, name, str(e))
            self.import_log += _("Erreur: Impossible de créer %s '%s': %s\n" % (model_name, name, str(e)))
            return False

    # --- LOGIQUE DE MAPPING ET IMPORTATION ---
    
    def _format_value(self, field, value):
        # ... (votre code de mapping inchangé) ...
        if value is None:
            return None
            
        value = str(value).strip().lower()

        selection_mapping = {
            'nature': {
                'livraison': 'livraison',
                'end to end': 'end_to_end',
                'services pro': 'services_pro',
                'all': 'all',
            },
            'bu': {
                'ict': 'ict',
                'cloud': 'cloud',
                'cybersecurity': 'cybersecurity',
                'formation': 'formation',
                'security': 'security',
            },
            'revenue_type': { 
                'recurrent': 'recurrent',
                'one shot': 'oneshot',
                'oneshot': 'oneshot',
            },
            'circuit': {
                'fast track': 'fast_track',
                'normal': 'normal',
            },
            'etat_projet': { 
                '0-annulé': 'annule',
                '1-non démarré': 'non_demarre',
                '2-en cours': 'en_cours_production',
                '3-en cours - provisionning': 'en_cours_provisioning',
                '4-en cours - livraison': 'en_cours_livraison',
                '5-terminé - pv/bl signé': 'termine_pv_bl_signe',
                '6-facturé - attente df': 'facture_attente_df',
                '7-cloturé': 'cloture',
                '8-suivi - contrat licence': 'suivi_contrat_licence',
                '8-suivi - contrat mixte': 'suivi_contrat_mixte',
                '8-suivi - contrat de services': 'suivi_contrat_services',
                '9-suspendu': 'suspendu',
                'cloturé': 'cloture',
                'non démarré': 'non_demarre',
                'en cours': 'en_cours_production',
                'terminé': 'termine_pv_bl_signe',
                'facturé': 'facture_attente_df',
                'draft': 'draft',
                'suspendu': 'suspendu',
            }
        }
        
        if field in selection_mapping:
            for excel_val, odoo_val in selection_mapping[field].items():
                if value == excel_val.lower():
                    return odoo_val
        
        fallback_values = {
            'nature': 'all',
            'bu': 'ict', 
            'domaine': 'others',
            'etat_projet': 'non_demarre',
            'revenue_type': 'oneshot',
            'circuit': 'normal'
        }
        
        return fallback_values.get(field, value)

    def _show_result_wizard(self):
        """Affiche le wizard avec les résultats"""
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'project.import.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
            'context': self.env.context,
        }

    def action_import_projects(self):
        """Logique principale d'importation des projets. (Nom de la méthode à utiliser dans le XML)"""
        self.import_log = ""
        self.success_count = 0
        self.error_count = 0
        self.created_users_count = 0
        self.created_partners_count = 0
        self.created_categories_count = 0
        
        Project = self.env['project.project'].sudo()

        if not openpyxl:
            raise UserError(_("Le module openpyxl n'est pas installé. Veuillez l'installer."))

        try:
            data = base64.b64decode(self.import_file)
            f = io.BytesIO(data)
            workbook = openpyxl.load_workbook(f)
            sheet = workbook.active
            
        except Exception as e:
            raise UserError(_("Erreur lors de la lecture du fichier : %s. Assurez-vous qu'il s'agit d'un fichier .xlsx valide." % str(e)))

        headers = [str(cell.value).strip() if cell.value is not None else '' for cell in sheet[1]]
        
        for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2):
            try:
                values = {}
                project_name = None
                partner_categories = []
                
                for excel_header, odoo_field in COLUMN_MAPPING.items():
                    if excel_header not in headers:
                        continue 
                        
                    col_index = headers.index(excel_header) 
                    cell_value = row[col_index].value
                    
                    if cell_value is None or str(cell_value).strip() == '':
                        continue
                        
                    # Champs de relation (Many2one)
                    if odoo_field in ['user_id', 'account_manager_id', 'presales_id']:
                        user_id = self._find_or_create_user(cell_value)
                        if user_id:
                            values[odoo_field] = user_id
                        
                    # Champs de relation (Many2one)
                    elif odoo_field == 'partner_id':
                        partner_id = self._find_or_create_partner(cell_value)
                        if partner_id:
                            values[odoo_field] = partner_id
                            
                    elif odoo_field == 'country_id':
                        country_id = self._find_or_create_misc('res.country', cell_value)
                        if country_id:
                            values[odoo_field] = country_id

                    # Champ de Catégorie Partenaire (M2M)
                    elif odoo_field == 'partner_category_ids':
                        category_id = self._find_or_create_misc('res.partner.category', cell_value)
                        if category_id:
                            partner_categories.append(category_id)
                            
                    # Champs de sélection
                    elif odoo_field in ['nature', 'bu', 'revenue_type', 'circuit', 'etat_projet', 'domaine']:
                        values[odoo_field] = self._format_value(odoo_field, cell_value)

                    # Champs de date
                    elif odoo_field in ['date_start', 'date_update']:
                        if isinstance(cell_value, datetime):
                            values[odoo_field] = cell_value.strftime('%Y-%m-%d %H:%M:%S')
                        elif isinstance(cell_value, str):
                             try:
                                 values[odoo_field] = datetime.strptime(cell_value.split()[0], '%Y-%m-%d').strftime('%Y-%m-%d')
                             except ValueError:
                                 pass 

                    # Nom du projet (clé)
                    elif odoo_field == 'name':
                        project_name = str(cell_value).strip()
                        values[odoo_field] = project_name
                        
                    # Champs numériques
                    elif odoo_field.startswith('cost_'):
                        try:
                            cleaned_value = re.sub(r'[^\d\.\,]', '', str(cell_value))
                            values[odoo_field] = float(cleaned_value.replace(',', '.') or 0)
                        except (ValueError, TypeError):
                            values[odoo_field] = 0.0
                            
                    # Autres champs simples
                    else:
                        values[odoo_field] = str(cell_value or '').strip()


                if not project_name:
                    raise UserError(_("Le nom du projet (colonne 'Nom') est manquant ou vide."))

                existing_project = Project.search([('name', '=', project_name)], limit=1)

                # Opérations sur le projet
                if existing_project and self.update_existing:
                    existing_project.write(values)
                    self.import_log += _("Ligne %d: Projet '%s' mis à jour.\n" % (row_index, project_name))
                    self.success_count += 1
                elif not existing_project and self.create_missing:
                    Project.create(values)
                    self.import_log += _("Ligne %d: Projet '%s' créé.\n" % (row_index, project_name))
                    self.success_count += 1
                
                # Mise à jour des catégories du partenaire après la création/mise à jour du projet
                if partner_categories and values.get('partner_id'):
                    partner = self.env['res.partner'].sudo().browse(values['partner_id'])
                    # Ajoute les nouvelles catégories sans écraser les existantes
                    current_categories = partner.category_id.ids
                    new_categories = list(set(current_categories + partner_categories))
                    partner.write({'category_id': [(6, 0, new_categories)]})
                
            except Exception as e:
                self.error_count += 1
                error_message = _("Erreur ligne %d pour projet '%s': %s" % (row_index, project_name or "N/A", str(e)))
                _logger.error(error_message)
                self.import_log += error_message + "\n"
                

        self.import_log += "\n--- RÉSUMÉ ---\n"
        self.import_log += _("Total Projets importés/mis à jour: %d\n" % self.success_count)
        self.import_log += _("Total Erreurs: %d\n" % self.error_count)
        self.import_log += _("Total Utilisateurs créés: %d\n" % self.created_users_count)

        return self._show_result_wizard()