# -*- coding: utf-8 -*-
# from odoo import http


# class Proptech/reportCustom(http.Controller):
#     @http.route('/proptech/report_custom/proptech/report_custom', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/proptech/report_custom/proptech/report_custom/objects', auth='public')
#     def list(self, **kw):
#         return http.request.render('proptech/report_custom.listing', {
#             'root': '/proptech/report_custom/proptech/report_custom',
#             'objects': http.request.env['proptech/report_custom.proptech/report_custom'].search([]),
#         })

#     @http.route('/proptech/report_custom/proptech/report_custom/objects/<model("proptech/report_custom.proptech/report_custom"):obj>', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('proptech/report_custom.object', {
#             'object': obj
#         })

