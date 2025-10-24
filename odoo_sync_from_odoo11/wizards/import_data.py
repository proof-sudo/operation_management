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
    'PM': 'user_id',
    'AM': 'am',
    'Presales': 'presales',
    'Nature': 'nature',
    'BU': 'bu',
    'Domaine': 'domaine',
    'Revenus': 'revenue_type',
    'Cat Recurrent': 'cat_recurrent',
    'Date IN': 'date_in',
    'Pays': 'pays',
    'Customer': 'partner_id',
    'Secteur': 'secteur',
    'Description du Projet': 'description',
    'Circuit': 'circuit',
    'SC': 'sc',
    'CAS Build': 'cas_build',
    'CAS Run': 'cas_run',
    'CAS Train': 'cas_train',
    'CAS Sw': 'cas_sw',
    'CAS Hw': 'cas_hw',
    'CAS': 'cas',
    'Statut': 'etat_projet',
}

# Domaine d'e-mail factice pour les utilisateurs créés
DEFAULT_USER_EMAIL_DOMAIN = 'neuronestech.com'


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
    created_users_count = fields.Integer(string='Utilisateurs créés', readonly=True)
    created_partners_count = fields.Integer(string='Clients créés', readonly=True)
    created_categories_count = fields.Integer(string='Catégories créées', readonly=True)

    # --- MÉTHODES DE GESTION DES ENREGISTREMENTS EXTERNES ---
    
    def _find_or_create_user(self, name):
        """ Recherche ou crée un utilisateur avec gestion des doublons et login unique """
        if not self.create_missing_records:
            return False
            
        name = str(name or '').strip()
        if not name or name.lower() in ('nan', 'none', 'default', 'n/a', 'na', ''):
            return False

        User = self.env['res.users'].sudo()
        
        # Recherche plus large pour éviter les doublons
        user = User.search([
            '|', ('name', '=ilike', name),
            '|', ('login', '=ilike', name),
            '|', ('email', '=ilike', name),
            ('partner_id.name', '=ilike', name)
        ], limit=1)
        
        if user:
            return user.id

        try:
            # Création du login de base (ex: Berenger ASSIELOU -> berenger.assielou)
            parts = re.findall(r'[a-zA-Z0-9]+', name.lower())
            if len(parts) > 1:
                login_base = ".".join(parts[:2])  # Prend seulement les 2 premières parties
            elif len(parts) == 1:
                login_base = parts[0]
            else:
                login_base = 'imported.user'
            
            # Nettoyer les caractères non alphanumériques sauf le point
            login_base = re.sub(r'[^a-z0-9\.]', '', login_base)
            
            if not login_base or len(login_base) < 3:
                login_base = 'imported.user'
                
            login_candidate = login_base
            login_suffix = 0
            
            # Garantir l'unicité du login
            while User.search([('login', '=', login_candidate)], limit=1):
                login_suffix += 1
                login_candidate = f"{login_base}.{login_suffix}"
            
            # CRÉATION CRITIQUE : Créer d'abord le partenaire
            Partner = self.env['res.partner'].sudo()
            partner_vals = {
                'name': name,
                'is_company': False,
                'company_type': 'person',
                'email': f'{login_candidate}@{DEFAULT_USER_EMAIL_DOMAIN}',
            }
            new_partner = Partner.create(partner_vals)
            
            # Puis créer l'utilisateur avec le partenaire
            user_vals = {
                'name': name,
                'login': login_candidate,
                'email': f'{login_candidate}@{DEFAULT_USER_EMAIL_DOMAIN}',
                'partner_id': new_partner.id,
                'company_id': self.env.company.id,
                'company_ids': [(6, 0, [self.env.company.id])],
                'notification_type': 'email',
                'groups_id': [(6, 0, [self.env.ref('base.group_user').id])]
            }
            
            new_user = User.create(user_vals)
            self.created_users_count += 1 
            _logger.info(f"Utilisateur créé: {name} (login: {login_candidate})")
            return new_user.id
            
        except Exception as e:
            _logger.error("Erreur recherche/création res.users '%s': %s", name, str(e))
            self.import_log += _("Erreur: Impossible de créer l'utilisateur '%s': %s\n" % (name, str(e)))
            return False

    def _find_or_create_partner(self, name):
        """ Recherche ou crée un partenaire (Client) avec gestion des doublons """
        if not self.create_missing_records:
            return False
            
        name = str(name or '').strip()
        if not name or name.lower() in ('nan', 'none', 'n/a', 'na', ''):
            return False
            
        Partner = self.env['res.partner'].sudo()
        
        # Recherche avec différentes stratégies pour éviter les doublons
        domain = [
            ('name', '=ilike', name),
            ('is_company', '=', True)
        ]
        partner = Partner.search(domain, limit=1)
        
        if not partner:
            # Recherche plus large
            partner = Partner.search([
                '|', ('name', '=ilike', name),
                ('commercial_company_name', '=ilike', name),
                ('is_company', '=', True)
            ], limit=1)
        
        if partner:
            return partner.id
            
        try:
            new_partner = Partner.create({
                'name': name,
                'is_company': True,
                'company_type': 'company',
            })
            self.created_partners_count += 1
            _logger.info(f"Partenaire créé: {name}")
            return new_partner.id
        except Exception as e:
            _logger.error("Erreur recherche/création res.partner '%s': %s", name, str(e))
            self.import_log += _("Erreur: Impossible de créer le partenaire '%s': %s\n" % (name, str(e)))
            return False

    def _find_or_create_misc(self, model_name, name, domain_filter=None):
        """ Recherche ou crée d'autres enregistrements avec gestion des doublons """
        if not self.create_missing_records:
            return False
            
        name = str(name or '').strip()
        if not name or name.lower() in ('nan', 'none', 'n/a', 'na', ''):
            return False

        Model = self.env[model_name].sudo()
        domain = [('name', '=ilike', name)]
        
        if domain_filter:
            domain.extend(domain_filter)
            
        if model_name == 'res.country':
            domain = ['|', ('name', '=ilike', name), ('code', '=ilike', name)]
            
        record = Model.search(domain, limit=1)
        if record:
            return record.id

        try:
            new_record = Model.create({'name': name})
            if model_name == 'res.partner.category':
                self.created_categories_count += 1
                _logger.info(f"Catégorie créée: {name}")
            return new_record.id
        except Exception as e:
            _logger.error("Erreur recherche/création %s '%s': %s", model_name, name, str(e))
            self.import_log += _("Erreur: Impossible de créer %s '%s': %s\n" % (model_name, name, str(e)))
            return False

    # --- LOGIQUE DE MAPPING ET IMPORTATION ---
    
    def _format_value(self, field, value):
        """ Formate les valeurs selon le type de champ """
        if value is None or str(value).strip().lower() in ('nan', 'none', 'n/a', 'na', ''):
            return None
            
        value_str = str(value).strip()

        selection_mapping = {
            'nature': {
                'livraison': 'livraison', 
                'end to end': 'end_to_end',
                'services pro': 'service_pro',
                'service pro': 'service_pro',
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
                'one-shot': 'oneshot',
            },
            'circuit': { 
                'fast track': 'fast',
                'fast': 'fast',
                'normal': 'normal', 
            },
            'domaine': {
                'datacenter facilities (dcf)': 'datacenter_facilities',
                'modern network integration (mni)': 'modern_network_integration',
                'agile infrastructure & cloud (aic)': 'agile_infrastructure_cloud',
                'business data integration (bdi)': 'business_data_integration',
                'digital workspace (dws)': 'digital_workspace',
                'secured it (sec)': 'secured_it',
                'expert & managed services - think': 'expert_managed_services_think',
                'expert & managed services - build': 'expert_managed_services_build',
                'expert & managed services - train': 'expert_managed_services_train',
                'expert & managed services - run': 'expert_managed_services_run',
                'none': 'none',
                'others': 'others',
            },
            'etat_projet': { 
                '0-annulé': 'cancelled', 
                '1-non démarré': 'non_demarre', 
                '2-en cours': 'en_cours_production', 
                '3-en cours - provisionning': 'en_cours_provisionning', 
                '4-en cours - livraison': 'en_cours_production',
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
                'cancelled': 'cancelled',
            }
        }
        
        if field in selection_mapping:
            for excel_val, odoo_val in selection_mapping[field].items():
                if value_str.lower() == excel_val.lower():
                    return odoo_val
        
        # Valeurs par défaut si non trouvé
        fallback_values = {
            'nature': 'all', 
            'bu': 'ict', 
            'domaine': 'others',
            'etat_projet': 'non_demarre', 
            'revenue_type': 'oneshot', 
            'circuit': 'normal'
        }
        
        return fallback_values.get(field, value_str)

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
        """Logique principale d'importation des projets."""
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
        _logger.info(f"En-têtes détectés: {headers}")
        
        for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=2):
            project_name = None
            try:
                values = {}
                
                for excel_header, odoo_field in COLUMN_MAPPING.items():
                    if excel_header not in headers:
                        continue 
                        
                    col_index = headers.index(excel_header) 
                    cell_value = row[col_index].value
                    
                    if cell_value is None or str(cell_value).strip() == '':
                        continue
                    
                    _logger.debug(f"Traitement: {excel_header} -> {odoo_field} = {cell_value}")
                    
                    # Champs Many2one sur res.users (PM, AM, Presales, SC)
                    if odoo_field in ['user_id', 'am', 'presales', 'sc']:
                        user_id = self._find_or_create_user(cell_value)
                        if user_id:
                            values[odoo_field] = user_id

                    # Many2one sur res.country (Pays)
                    elif odoo_field == 'pays':
                        country_id = self._find_or_create_misc('res.country', cell_value)
                        if country_id:
                            values[odoo_field] = country_id

                    # Champ date (Date IN)
                    elif odoo_field == 'date_in':
                        if isinstance(cell_value, datetime):
                            values[odoo_field] = cell_value.strftime('%Y-%m-%d')
                        elif isinstance(cell_value, str):
                            try:
                                # Essayer différents formats de date
                                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%m/%d/%Y']:
                                    try:
                                        values[odoo_field] = datetime.strptime(cell_value, fmt).strftime('%Y-%m-%d')
                                        break
                                    except ValueError:
                                        continue
                            except ValueError:
                                _logger.warning(f"Format de date invalide pour {cell_value}")

                    # Many2one sur res.partner (Customer)
                    elif odoo_field == 'partner_id':
                        partner_id = self._find_or_create_partner(cell_value)
                        if partner_id:
                            values[odoo_field] = partner_id
                            
                    # Many2one sur res.partner.category (Secteur)
                    elif odoo_field == 'secteur':
                        category_id = self._find_or_create_misc('res.partner.category', cell_value)
                        if category_id:
                            values[odoo_field] = category_id
                            
                    # Champs de sélection
                    elif odoo_field in ['nature', 'bu', 'domaine', 'revenue_type', 'circuit', 'etat_projet']:
                        formatted_value = self._format_value(odoo_field, cell_value)
                        if formatted_value:
                            values[odoo_field] = formatted_value
                            
                    # Champs monétaires (CAS)
                    elif odoo_field in ['cas_build', 'cas_run', 'cas_train', 'cas_sw', 'cas_hw', 'cas']:
                        try:
                            if isinstance(cell_value, (int, float)):
                                values[odoo_field] = float(cell_value)
                            elif isinstance(cell_value, str):
                                cleaned_value = re.sub(r'[^\d\.\,]', '', str(cell_value))
                                values[odoo_field] = float(cleaned_value.replace(',', '.') or 0)
                            else:
                                values[odoo_field] = 0.0
                        except (ValueError, TypeError) as e:
                            _logger.warning(f"Valeur monétaire invalide pour {cell_value}: {e}")
                            values[odoo_field] = 0.0
                            
                    # Champ texte (Nom)
                    elif odoo_field == 'name':
                        project_name = str(cell_value).strip()
                        values[odoo_field] = project_name
                        
                    # Champs de texte simples
                    elif odoo_field in ['description', 'cat_recurrent']:
                        values[odoo_field] = str(cell_value).strip()

                if not project_name:
                    self.error_count += 1
                    self.import_log += _("Ligne %d: Nom du projet manquant, ligne ignorée.\n" % row_index)
                    continue

                # Recherche du projet existant
                existing_project = Project.search([('name', '=', project_name)], limit=1)

                # Opérations sur le projet
                if existing_project and self.update_existing:
                    existing_project.write(values)
                    self.import_log += _("Ligne %d: Projet '%s' mis à jour.\n" % (row_index, project_name))
                    self.success_count += 1
                    _logger.info(f"Projet mis à jour: {project_name}")
                    
                elif not existing_project and self.create_missing:
                    new_project = Project.create(values)
                    self.import_log += _("Ligne %d: Projet '%s' créé.\n" % (row_index, project_name))
                    self.success_count += 1
                    _logger.info(f"Projet créé: {project_name} (ID: {new_project.id})")
                else:
                    self.import_log += _("Ligne %d: Projet '%s' ignoré (existe déjà et mise à jour désactivée).\n" % (row_index, project_name))
                
            except Exception as e:
                self.error_count += 1
                error_message = _("Erreur ligne %d pour projet '%s': %s" % (row_index, project_name or "N/A", str(e)))
                _logger.error(error_message)
                self.import_log += error_message + "\n"
                
        # Résumé final
        self.import_log += "\n--- RÉSUMÉ ---\n"
        self.import_log += _("Total Projets importés/mis à jour: %d\n" % self.success_count)
        self.import_log += _("Total Erreurs: %d\n" % self.error_count)
        self.import_log += _("Total Utilisateurs créés: %d\n" % self.created_users_count)
        self.import_log += _("Total Clients créés: %d\n" % self.created_partners_count)
        self.import_log += _("Total Catégories créées: %d\n" % self.created_categories_count)

        _logger.info(f"Import terminé: {self.success_count} succès, {self.error_count} erreurs")

        return self._show_result_wizard()