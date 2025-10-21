from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
import xlrd
import logging

_logger = logging.getLogger(__name__)

class ProjectImportWizard(models.TransientModel):
    _name = 'project.import.wizard'
    _description = 'Wizard d\'import de projets depuis Excel'

    import_file = fields.Binary(
        string='Fichier Excel',
        required=True,
        help='Fichier Excel (.xlsx) contenant les données des projets'
    )
    import_filename = fields.Char(string='Nom du fichier')
    
    # Options d'import
    update_existing = fields.Boolean(
        string='Mettre à jour les projets existants',
        default=True,
        help='Si cochée, met à jour les projets existants basés sur le nom'
    )
    create_missing = fields.Boolean(
        string='Créer les projets manquants',
        default=True,
        help='Si cochée, crée les projets qui n\'existent pas'
    )
    
    # Résultats de l'import
    import_log = fields.Text(
        string='Journal d\'import',
        readonly=True
    )
    success_count = fields.Integer(
        string='Projets importés/mis à jour',
        readonly=True
    )
    error_count = fields.Integer(
        string='Erreurs',
        readonly=True
    )

    def action_import(self):
        """Action principale d'import du fichier Excel"""
        self.ensure_one()
        
        if not self.import_file:
            raise UserError(_("Veuillez sélectionner un fichier Excel à importer."))

        log_messages = ["=== DÉBUT DE L'IMPORT EXCEL ==="]
        success_count = 0
        error_count = 0

        try:
            # Décoder le fichier
            file_content = base64.b64decode(self.import_file)
            workbook = xlrd.open_workbook(file_contents=file_content)
            sheet = workbook.sheet_by_index(0)  # Première feuille
            
            # Lire les en-têtes
            headers = [header.lower().strip() for header in sheet.row_values(0)]
            required_headers = ['name', 'nature', 'domaine', 'secteur']
            
            # Vérifier les en-têtes requis
            for req_header in required_headers:
                if req_header not in headers:
                    raise UserError(_(
                        f"Colonne requise manquante: {req_header}. "
                        f"Colonnes trouvées: {headers}"
                    ))

            # Mapping des colonnes
            col_mapping = {
                'name': headers.index('name'),
                'nature': headers.index('nature') if 'nature' in headers else None,
                'domaine': headers.index('domaine') if 'domaine' in headers else None,
                'secteur': headers.index('secteur') if 'secteur' in headers else None,
                'bu': headers.index('bu') if 'bu' in headers else None,
                'circuit': headers.index('circuit') if 'circuit' in headers else None,
                'priorite': headers.index('priorite') if 'priorite' in headers else None,
                'etat_projet': headers.index('etat_projet') if 'etat_projet' in headers else None,
                'cas': headers.index('cas') if 'cas' in headers else None,
                'cafy': headers.index('cafy') if 'cafy' in headers else None,
                'am': headers.index('am') if 'am' in headers else None,
                'presales': headers.index('presales') if 'presales' in headers else None,
            }

            # Parcourir les lignes de données
            for row_idx in range(1, sheet.nrows):
                try:
                    row_data = sheet.row_values(row_idx)
                    project_vals = self._prepare_project_vals(row_data, col_mapping)
                    
                    # Rechercher le projet existant
                    project = self.env['project.project'].search([
                        ('name', '=', project_vals.get('name'))
                    ], limit=1)
                    
                    if project and self.update_existing:
                        # Mettre à jour le projet existant
                        project.write(project_vals)
                        log_messages.append(f"✓ Ligne {row_idx + 1}: Projet '{project_vals['name']}' mis à jour")
                        success_count += 1
                        
                    elif self.create_missing:
                        # Créer un nouveau projet
                        self.env['project.project'].create(project_vals)
                        log_messages.append(f"✓ Ligne {row_idx + 1}: Projet '{project_vals['name']}' créé")
                        success_count += 1
                    else:
                        log_messages.append(f"⏭ Ligne {row_idx + 1}: Projet '{project_vals['name']}' ignoré (création désactivée)")
                        
                except Exception as e:
                    error_count += 1
                    error_msg = f"✗ Ligne {row_idx + 1}: Erreur - {str(e)}"
                    log_messages.append(error_msg)
                    _logger.error(error_msg)

            # Finaliser le journal
            log_messages.append(f"\n=== RÉSUMÉ ===")
            log_messages.append(f"Projets traités avec succès: {success_count}")
            log_messages.append(f"Erreurs: {error_count}")
            log_messages.append("=== FIN DE L'IMPORT ===")

            # Mettre à jour le wizard avec les résultats
            self.write({
                'import_log': '\n'.join(log_messages),
                'success_count': success_count,
                'error_count': error_count
            })

            # Afficher un message de confirmation
            if error_count == 0:
                return {
                    'type': 'ir.actions.act_window',
                    'res_model': self._name,
                    'res_id': self.id,
                    'view_mode': 'form',
                    'target': 'new',
                    'context': {'import_success': True}
                }
            else:
                warning_msg = _(
                    f"Import terminé avec {success_count} succès et {error_count} erreurs. "
                    f"Consultez le journal pour plus de détails."
                )
                raise UserError(warning_msg)

        except Exception as e:
            _logger.error(f"Erreur lors de l'import Excel: {str(e)}")
            raise UserError(_(f"Erreur lors de la lecture du fichier Excel: {str(e)}"))

    def _prepare_project_vals(self, row_data, col_mapping):
        """Prépare les valeurs du projet à partir des données Excel"""
        vals = {}
        
        # Champs de base
        if col_mapping['name'] is not None:
            vals['name'] = str(row_data[col_mapping['name']]).strip()
        
        # Champs de sélection
        selection_fields = {
            'nature': 'nature',
            'domaine': 'domaine', 
            'bu': 'bu',
            'circuit': 'circuit',
            'priorite': 'priorite',
            'etat_projet': 'etat_projet'
        }
        
        for excel_col, odoo_field in selection_fields.items():
            if col_mapping[excel_col] is not None:
                cell_value = str(row_data[col_mapping[excel_col]]).strip()
                if cell_value:
                    vals[odoo_field] = cell_value
        
        # Champs numériques
        numeric_fields = {'cas': 'cas', 'cafy': 'cafy'}
        for excel_col, odoo_field in numeric_fields.items():
            if col_mapping[excel_col] is not None:
                try:
                    cell_value = row_data[col_mapping[excel_col]]
                    if cell_value:
                        vals[odoo_field] = float(cell_value)
                except (ValueError, TypeError):
                    pass  # Garde la valeur par défaut si conversion impossible
        
        # Gestion du secteur (Many2one)
        if col_mapping['secteur'] is not None:
            secteur_name = str(row_data[col_mapping['secteur']]).strip()
            if secteur_name:
                secteur = self.env['res.partner.category'].search([
                    ('name', '=ilike', secteur_name)
                ], limit=1)
                if secteur:
                    vals['secteur'] = secteur.id
                else:
                    _logger.warning(f"Secteur non trouvé: {secteur_name}")
        
        # Gestion des utilisateurs (AM, Presales)
        user_fields = {'am': 'am', 'presales': 'presales'}
        for excel_col, odoo_field in user_fields.items():
            if col_mapping[excel_col] is not None:
                user_login = str(row_data[col_mapping[excel_col]]).strip()
                if user_login:
                    user = self.env['res.users'].search([
                        '|', ('login', '=ilike', user_login),
                        ('name', '=ilike', user_login)
                    ], limit=1)
                    if user:
                        vals[odoo_field] = user.id
                    else:
                        _logger.warning(f"Utilisateur non trouvé: {user_login}")
        
        return vals

    def action_show_log(self):
        """Action pour afficher le journal d'import"""
        return {
            'type': 'ir.actions.act_window',
            'name': 'Journal d\'import',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
            'context': {'show_log': True}
        }