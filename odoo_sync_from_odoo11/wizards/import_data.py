from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
import io
import logging
from datetime import datetime

_logger = logging.getLogger(__name__)

try:
    import openpyxl
except ImportError:
    _logger.warning("Le module openpyxl n'est pas installé. Installation requise: pip install openpyxl")
    openpyxl = None

class ProjectImportWizard(models.TransientModel):
    _name = 'project.import.wizard'
    _description = 'Wizard d\'import de projets depuis Excel'

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
    
    import_log = fields.Text(string='Journal d\'import', readonly=True)
    success_count = fields.Integer(string='Succès', readonly=True)
    error_count = fields.Integer(string='Erreurs', readonly=True)

    def action_import(self):
        """Action principale d'import"""
        if not self.import_file:
            raise UserError(_("Veuillez sélectionner un fichier Excel."))

        if not openpyxl:
            raise UserError(_(
                "Le module 'openpyxl' n'est pas installé. "
                "Veuillez l'installer avec: pip install openpyxl"
            ))

        log_messages = ["=== DÉBUT DE L'IMPORT ==="]
        success_count = 0
        error_count = 0

        try:
            # Lecture du fichier avec openpyxl
            file_content = base64.b64decode(self.import_file)
            workbook = openpyxl.load_workbook(filename=io.BytesIO(file_content), data_only=True)
            sheet = workbook.active  # Première feuille
            
            # Lecture des en-têtes
            headers = []
            for cell in sheet[1]:  # Première ligne
                headers.append(str(cell.value).strip() if cell.value else "")
            
            _logger.info(f"En-têtes détectés: {headers}")
            
            # Mapping des colonnes
            col_mapping = self._create_column_mapping(headers)
            
            # Vérifier les colonnes obligatoires
            if col_mapping['name'] is None:
                raise UserError(_("La colonne 'Nom' est obligatoire dans le fichier Excel."))

            # Parcourir les lignes de données
            for row_idx in range(2, sheet.max_row + 1):  # Commence à la ligne 2
                try:
                    row_data = []
                    for col_idx in range(1, sheet.max_column + 1):
                        cell_value = sheet.cell(row=row_idx, column=col_idx).value
                        row_data.append(cell_value)
                    
                    project_vals = self._prepare_project_vals(row_data, col_mapping, row_idx)
                    
                    project_name = project_vals.get('name')
                    if not project_name:
                        log_messages.append(f"⚠️ Ligne {row_idx}: Nom manquant, ligne ignorée")
                        continue

                    project = self.env['project.project'].search([
                        ('name', '=', project_name)
                    ], limit=1)
                    
                    if project and self.update_existing:
                        project.write(project_vals)
                        log_messages.append(f"✓ Ligne {row_idx}: '{project_name}' mis à jour")
                        success_count += 1
                    elif self.create_missing:
                        self.env['project.project'].create(project_vals)
                        log_messages.append(f"✓ Ligne {row_idx}: '{project_name}' créé")
                        success_count += 1
                    else:
                        log_messages.append(f"⏭ Ligne {row_idx}: '{project_name}' ignoré")
                        
                except Exception as e:
                    error_count += 1
                    log_messages.append(f"✗ Ligne {row_idx}: Erreur - {str(e)}")
                    _logger.error(f"Erreur ligne {row_idx}: {str(e)}")

            log_messages.append(f"\n=== RÉSUMÉ ===")
            log_messages.append(f"Projets traités avec succès: {success_count}")
            log_messages.append(f"Erreurs: {error_count}")
            log_messages.append("=== FIN DE L'IMPORT ===")

            self.write({
                'import_log': '\n'.join(log_messages),
                'success_count': success_count,
                'error_count': error_count
            })

            return self._show_result_wizard()

        except Exception as e:
            _logger.error(f"Erreur générale import: {str(e)}")
            raise UserError(_(f"Erreur lors de l'import: {str(e)}"))

    def _create_column_mapping(self, headers):
        """Crée le mapping entre les colonnes Excel et les champs Odoo"""
        mapping = {}
        
        # Mapping direct des colonnes
        column_mapping = {
            'Nom': 'name',
            'Nature': 'nature',
            'BU': 'bu',
            'Domaine': 'domaine',
            'Revenus': 'revenue_type',
            'Cat Recurrent': 'cat_recurrent',
            'AM': 'am',
            'Presales': 'presales',
            'Date IN': 'date_in',
            'Pays': 'pays',
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
            'Update Date': 'write_date',
            'PM': 'user_id',  # Chef de projet
            'Customer': 'partner_id'  # Client
        }
        
        for excel_col, odoo_field in column_mapping.items():
            if excel_col in headers:
                mapping[odoo_field] = headers.index(excel_col)
            else:
                mapping[odoo_field] = None
                _logger.warning(f"Colonne '{excel_col}' non trouvée dans le fichier")
        
        return mapping

    def _prepare_project_vals(self, row_data, col_mapping, row_num):
        """Prépare les valeurs pour la création/mise à jour du projet"""
        vals = {}
        
        # Champ nom (obligatoire)
        if col_mapping['name'] is not None:
            cell_value = row_data[col_mapping['name']]
            if cell_value is not None:
                vals['name'] = str(cell_value).strip()
        
        # Champs de sélection
        selection_fields = {
            'nature': 'nature',
            'bu': 'bu',
            'domaine': 'domaine',
            'revenue_type': 'revenue_type',
            'circuit': 'circuit',
            'etat_projet': 'etat_projet'
        }
        
        for odoo_field in selection_fields:
            if col_mapping[odoo_field] is not None:
                cell_value = row_data[col_mapping[odoo_field]]
                if cell_value is not None:
                    cell_value_str = str(cell_value).strip()
                    if cell_value_str and cell_value_str.lower() != 'none' and cell_value_str != '':
                        converted_value = self._convert_selection_value(odoo_field, cell_value_str)
                        if converted_value:
                            vals[odoo_field] = converted_value
        
        # Champs texte
        text_fields = ['cat_recurrent', 'description']
        for field in text_fields:
            if col_mapping.get(field) is not None:
                cell_value = row_data[col_mapping[field]]
                if cell_value is not None:
                    cell_value_str = str(cell_value).strip()
                    if cell_value_str:
                        vals[field] = cell_value_str
        
        # Champs numériques
        numeric_fields = {
            'cas_build': 'cas_build',
            'cas_run': 'cas_run', 
            'cas_train': 'cas_train',
            'cas_sw': 'cas_sw',
            'cas_hw': 'cas_hw',
            'cas': 'cas'
        }
        
        for excel_field, odoo_field in numeric_fields.items():
            if col_mapping[odoo_field] is not None:
                cell_value = row_data[col_mapping[odoo_field]]
                if cell_value is not None:
                    try:
                        if cell_value != '':
                            vals[odoo_field] = float(cell_value)
                    except (ValueError, TypeError) as e:
                        _logger.warning(f"Ligne {row_num}: Valeur numérique invalide pour {excel_field}: {cell_value}")
        
        # Champs dates
        date_fields = {'date_in': 'date_in'}
        for odoo_field in date_fields:
            if col_mapping[odoo_field] is not None:
                cell_value = row_data[col_mapping[odoo_field]]
                if cell_value:
                    try:
                        if isinstance(cell_value, datetime):
                            # C'est déjà un objet datetime
                            vals[odoo_field] = cell_value.strftime('%Y-%m-%d')
                        elif isinstance(cell_value, str):
                            # Tentative de parsing de date string
                            date_str = cell_value.strip()
                            if date_str:
                                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y']:
                                    try:
                                        date_obj = datetime.strptime(date_str, fmt)
                                        vals[odoo_field] = date_obj.strftime('%Y-%m-%d')
                                        break
                                    except ValueError:
                                        continue
                    except Exception as e:
                        _logger.warning(f"Ligne {row_num}: Date invalide pour {odoo_field}: {cell_value}")
        
        # Champs relationnels (Many2one)
        relational_fields = {
            'secteur': ('res.partner.category', 'secteur'),
            'pays': ('res.country', 'pays'),
            'user_id': ('res.users', 'user_id'),  # PM - Chef de projet
            'am': ('res.users', 'am'),
            'presales': ('res.users', 'presales'),
            'sc': ('res.users', 'sc'),
            'partner_id': ('res.partner', 'partner_id')  # Customer - Client
        }
        
        for odoo_field, (model, field_name) in relational_fields.items():
            if col_mapping[odoo_field] is not None:
                cell_value = row_data[col_mapping[odoo_field]]
                if cell_value is not None:
                    cell_value_str = str(cell_value).strip()
                    if cell_value_str:
                        record = self._find_related_record(model, cell_value_str, row_num)
                        if record:
                            vals[field_name] = record.id
                        else:
                            _logger.warning(f"Ligne {row_num}: {model} non trouvé: {cell_value_str}")

        return vals

    def _convert_selection_value(self, field, value):
        """Convertit les valeurs de sélection depuis Excel vers Odoo"""
        value = value.strip().lower()
        
        # Mapping des valeurs de sélection
        selection_mapping = {
            'nature': {
                'all': 'all',
                'end to end': 'end_to_end',
                'livraison': 'livraison',
                'service pro': 'service_pro',
            },
            'revenue_type': {
                'one shot': 'oneshot',
                'oneshot': 'oneshot',
                'recurrent': 'recurrent',
            },
            'circuit': {
                'fast': 'fast',
                'fast track': 'fast',
                'normal': 'normal',
            },
            'bu': {
                'ict': 'ict',
                'cloud': 'cloud',
                'cybersecurity': 'cybersecurity',
                'formation': 'formation',
                'security': 'security',
            }
        }
        
        if field in selection_mapping:
            for excel_val, odoo_val in selection_mapping[field].items():
                if value == excel_val.lower():
                    return odoo_val
        
        return value

    def _find_related_record(self, model, search_value, row_num):
        """Trouve un enregistrement related par nom ou autre champ"""
        try:
            if model == 'res.users':
                return self.env[model].search([
                    '|', ('name', '=ilike', search_value),
                    ('login', '=ilike', search_value)
                ], limit=1)
            elif model == 'res.partner':
                return self.env[model].search([
                    ('name', '=ilike', search_value)
                ], limit=1)
            elif model == 'res.partner.category':
                return self.env[model].search([
                    ('name', '=ilike', search_value)
                ], limit=1)
            elif model == 'res.country':
                return self.env[model].search([
                    '|', ('name', '=ilike', search_value),
                    ('code', '=ilike', search_value)
                ], limit=1)
        except Exception as e:
            _logger.error(f"Erreur recherche {model} '{search_value}': {str(e)}")
        return None

    def _show_result_wizard(self):
        """Affiche le wizard avec les résultats"""
        return {
            'type': 'ir.actions.act_window',
            'name': 'Résultat de l\'import',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }