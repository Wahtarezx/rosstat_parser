import zipfile
import os
from io import BytesIO
from django.contrib import admin, messages

from django.http import HttpResponse
from django.urls import path
from django.shortcuts import redirect
from apps.rosstat_parser.models import Region


@admin.register(Region)
class RegionAdmin(admin.ModelAdmin):
    list_display = ['id', 'name']
    ordering = ['name']
    change_list_template = "region_changelist.html"
    actions = ['download_selected_tables']

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                "download-all-tables/",
                self.admin_site.admin_view(self.download_all_tables),
                name="rosstat_parser_region_download-all-tables"
            )
        ]
        return custom_urls + urls

    def download_all_tables(self, request):
        buffer = BytesIO()
        missing = []
        with zipfile.ZipFile(buffer, "w") as zf:
            for region in Region.objects.all():
                if not region.analytical_table:
                    continue
                try:
                    file_path = region.analytical_table.path
                except (ValueError, NotImplementedError):
                    missing.append(region.name)
                    continue
                if not os.path.exists(file_path):
                    missing.append(region.name)
                    continue
                filename = f"{region.name}_{os.path.basename(file_path)}"
                zf.write(file_path, filename)

        if buffer.tell() == 0:
            messages.error(request, "Нет доступных таблиц для скачивания")
            return redirect("..")

        if missing:
            messages.warning(
                request,
                "Файлы отсутствуют на диске для регионов: " + ", ".join(missing),
            )

        buffer.seek(0)
        response = HttpResponse(buffer, content_type="application/zip")
        response["Content-Disposition"] = "attachment; filename=all_region_tables.zip"
        return response

    def download_selected_tables(self, request, queryset):
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as zf:
            for region in queryset:
                if region.analytical_table:
                    try:
                        file_path = region.analytical_table.path
                        filename = f"{region.name}_{os.path.basename(file_path)}"
                        zf.write(file_path, filename)
                    except Exception as e:
                        continue

        if buffer.tell() == 0:
            messages.error(request, "Нет доступных таблиц для скачивания среди выбранных регионов")
            return redirect("..")

        buffer.seek(0)
        response = HttpResponse(buffer, content_type="application/zip")
        response["Content-Disposition"] = "attachment; filename=selected_region_tables.zip"
        return response

    download_selected_tables.short_description = "Скачать таблицы выбранных регионов"
