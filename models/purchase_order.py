# -*- coding: utf-8 -*-
from odoo import models, fields, api

class PurchaseOrder(models.Model):
    _inherit = 'purchase.order'

    def action_open_racking_report_wizard(self):
        self.ensure_one()
        wiz = self.env['wb.racking.report.wizard'].create({
            'purchase_id': self.id,
        })
        return wiz.action_generate_excel_report()
