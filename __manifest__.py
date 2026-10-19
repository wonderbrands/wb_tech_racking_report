# -*- coding: utf-8 -*-
{
    'name': 'WB Tech Racking Report',
    'summary': 'Reporte detallado de movimientos de rackeo (STOR) en Excel por PO',
    'description': """
        Módulo para generar reportes en Excel (.xlsx) por Orden de Compra (PO).
        Analiza todos los movimientos STOR, descompone producto por producto en pestañas individuales,
        mapea origen, destino, usuario, estado de automatización y balance de rackeo.
    """,
    'version': '18.0.1.0.0',
    'category': 'Inventory/Warehouse',
    'author': 'Wonderbrands',
    'website': 'https://www.wonderbrands.co',
    'license': 'LGPL-3',
    'depends': [
        'wmds',
        'purchase',
        'stock',
    ],
    'external_dependencies': {
        'python': [
            'openpyxl',
        ],
    },
    'data': [
        'security/ir.model.access.csv',
        'wizard/racking_report_wizard_views.xml',
        'views/purchase_order_views.xml',
    ],
    'installable': True,
    'application': False,
}
