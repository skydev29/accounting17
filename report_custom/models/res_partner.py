from odoo import models, fields, api


class projectProject(models.Model):
    _inherit = 'project.project'

    use_documents = fields.Boolean(string="Is Coa Installed")


class accountMove(models.Model):
    _inherit = 'account.move'

    tax_closing_end_date = fields.Date(string="Is Coa Installed")


class ResPartner(models.Model):
    _inherit = 'res.partner'

    is_coa_installed = fields.Boolean(string="Is Coa Installed")

