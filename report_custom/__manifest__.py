# -*- coding: utf-8 -*-
{
    'name': "report_custom",

    'summary': "Short (1 phrase/line) summary of the module's purpose",

    'description': """
Long description of module's purpose
    """,

    'author': "My Company",
    'website': "https://www.yourcompany.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/15.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '0.1',

    #hh hh

    # any module necessary for this one to work correctly
    'depends': ['base','web','l10n_gcc_invoice','l10n_sa', 'sale'],

    # always loaded
    'data': [
        # 'security/ir.model.access.csv',
        'views/report_invoice.xml',
        'views/report_sale_order.xml',
        'views/report_templates.xml',
        'views/views.xml',
        'views/templates.xml',
    ],
    'assets': {

            'web.report_assets_common': [
                'report_custom/static/src/css/style.css',
                'report_custom/static/src/css/report.scss',
            ],

        },

}

