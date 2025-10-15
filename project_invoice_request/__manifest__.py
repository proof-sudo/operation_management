{
    'name': 'Demande de Facturation Projet',
    'version': '18.0',
    'category': 'Project',
    'sequence': 3,
    'summary': 'Workflow de validation pour les demandes de facturation sur les projets',
    'description': """
        Ajoute un workflow de validation pour les demandes de facturation:
        - Bouton sur le formulaire du projet
        - Wizard pour sélectionner les lignes et quantités à facturer
        - Gestion des facturations partielles
        - Validation en plusieurs étapes
    """,
    'author': 'Votre Société',
    'depends': ['project', 'sale_project', 'account','sale','mail'],
    'data': [
        'security/ir.model.access.csv',
        'security/security.xml',
        'data/sequence_data.xml',
        'data/mail_template.xml',
        'views/project_views.xml',
        'views/invoice_request_views.xml',
        # 'views/kaban_heritage.xml',
        'wizard/invoice_request_wizard_views.xml',
    ],
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
