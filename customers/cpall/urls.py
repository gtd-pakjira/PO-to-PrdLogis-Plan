from django.urls import path

from . import views

app_name = "cpall"

urlpatterns = [
    path("", views.index, name="index"),
    path("po/", views.po_list, name="po_list"),
    path("plans/", views.plan_list, name="plan_list"),
    path("import/", views.import_form, name="import_form"),
    path("import/submit/", views.import_submit, name="import_submit"),
    path("import/confirm-duplicates/", views.confirm_duplicates, name="confirm_duplicates"),
    path("po/<int:po_import_id>/resolve-locations/", views.resolve_locations, name="resolve_locations"),
    path("po/<int:po_import_id>/resolve-products/", views.resolve_products, name="resolve_products"),
    path("plan/new/", views.new_plan_submit, name="new_plan_submit"),
    path("plan/buffer/", views.buffer_form, name="buffer_form"),
    path("plan/buffer/submit/", views.buffer_form_submit, name="buffer_form_submit"),
    path("plan/<int:plan_run_id>/", views.view_plan, name="view_plan"),
    path("plan/<int:plan_run_id>/download/production/", views.download_production, name="download_production"),
    path("plan/<int:plan_run_id>/download/logistic/<str:group_name>/", views.download_logistic, name="download_logistic"),
    path("plan/<int:plan_run_id>/download/all/", views.download_all_zip, name="download_all_zip"),
    path("templates/", views.template_list, name="template_list"),
    path("templates/<str:key>/download/", views.template_download, name="template_download"),
    path("templates/<str:key>/upload/", views.template_upload, name="template_upload"),
    path("templates/<str:key>/versions/", views.template_versions, name="template_versions"),
    path("templates/<str:key>/versions/<int:version_id>/restore/", views.template_version_restore,
         name="template_version_restore"),
    path("templates/<str:key>/versions/<int:version_id>/delete/", views.template_version_delete,
         name="template_version_delete"),
    path("templates/<str:key>/view/", views.template_view, name="template_view"),
    path("plan/<int:plan_run_id>/table/production/", views.view_production_table, name="view_production_table"),
    path("plan/<int:plan_run_id>/table/logistic/<str:group_name>/", views.view_logistic_table, name="view_logistic_table"),
    path("po/<int:po_import_id>/delete/", views.delete_po_import_view, name="delete_po_import"),
    path("po/<int:po_import_id>/", views.view_po_detail, name="view_po_detail"),
    path("po/<int:po_import_id>/download/", views.download_po, name="download_po"),
    path("plan/<int:plan_run_id>/delete/", views.delete_plan_run_view, name="delete_plan_run"),
]
