from django.contrib import admin

from .models import LocationMapping, SkuMaster


@admin.register(SkuMaster)
class SkuMasterAdmin(admin.ModelAdmin):
    list_display = ("barcode", "name_th", "pack_size", "unit_price", "updated_at")
    search_fields = ("barcode", "name_th", "name_en")


@admin.register(LocationMapping)
class LocationMappingAdmin(admin.ModelAdmin):
    list_display = ("fc_code", "name_th", "group", "sub_location", "updated_at")
    list_filter = ("group",)
    search_fields = ("fc_code", "name_th", "sub_location")
