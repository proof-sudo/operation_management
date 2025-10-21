# -*- coding: utf-8 -*-
{
    'name': 'Odoo17 Sync Receiver',
    'version': '1.0',
    'sequence': 3,
    'summary': 'Receives Sale Orders and Invoices from Odoo11 and syncs data',
    'description': """
        Module de réception et synchronisation des données depuis Odoo11.
        - Crée les clients si inexistants
        - Crée les produits si inexistants
        - Associe les entrepôts
        - Utilise l\'admin si l\'utilisateur n\'existe pas
    """,
    'author': 'Djakaridja Traore',
    'category': 'Sales',
    'website': 'https://yourcompany.com',
    'depends': ['base', 'sale', 'stock', 'account','sale_management',  # Pour hériter de sale.order
        'project','purchase'],  # Pour hériter de account.move
    'license': 'LGPL-3',
    'data': [
        'security/ir.model.access.csv',
        # 'views/res_partner_view.xml',
        'wizards/import_wizard_views.xml',
        'views/projet_inherit_view.xml',# Fichier de sécurité (très important !)
        'views/sale_order_view.xml',
        'views/project_list_view_inherit.xml',
        'views/create_project_wizard_view.xml',
        'views/purchase_order_view.xml',
       
        
        
        # ici on pourrait ajouter des vues ou des sécurités si nécessaire
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
}
