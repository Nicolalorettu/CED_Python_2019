from django.contrib import admin
from .models import tools_credentials, sites_tracker, order_tracker, tkt_money_value, discount_tracker, \
    contract_tracker, kpi_tracker, supervisor_tracker


class tools_credentialsadmin(admin.ModelAdmin):
    list_display = ("user", "password", "description")

class sites_trackeradmin(admin.ModelAdmin):
    list_display = ("sitename", "manager", "fromdate", "todate")

class contract_trackeradmin(admin.ModelAdmin):
    list_display = ("name", "order", "fromdate", "todate")

class kpi_trackeradmin(admin.ModelAdmin):
    list_display = ("contract", "order", "name", "description", "target", "tier1", "sortorder")
    list_filter = ("contract", "order")

class order_trackeradmin(admin.ModelAdmin):
    list_display = ("pafdesc", "description", "site")

class supervisor_trackeradmin(admin.ModelAdmin):
    list_display = ( "surname", "name", "order", "fromdate", "todate")

class tkt_money_valueadmin(admin.ModelAdmin):
    list_display = ("contract", "order", "pafdesc", "description", "value", "site")
    list_filter = ("contract", "order", "site")

class discount_trackeradmin(admin.ModelAdmin):
    list_display = ("description", "disc_value", "fromdate", "todate")


admin.site.register(tools_credentials, tools_credentialsadmin)
admin.site.register(tkt_money_value, tkt_money_valueadmin)
admin.site.register(sites_tracker, sites_trackeradmin)
admin.site.register(order_tracker, order_trackeradmin)
admin.site.register(discount_tracker, discount_trackeradmin)
admin.site.register(contract_tracker, contract_trackeradmin)
admin.site.register(kpi_tracker, kpi_trackeradmin)
admin.site.register(supervisor_tracker, supervisor_trackeradmin)
