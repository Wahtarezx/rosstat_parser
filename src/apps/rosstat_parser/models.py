from django.db import models
from django.core.files.storage import FileSystemStorage


class OriginalNameStorage(FileSystemStorage):
    def get_available_name(self, name, max_length=None):
        return name

    def _save(self, name, content):
        # Если файл уже существует, удаляем его
        if self.exists(name):
            self.delete(name)
        return super()._save(name, content)


def region_upload_path(instance, filename):
    return f"regions/{instance.name}/tables/{filename}"


class Region(models.Model):
    name = models.CharField(
        max_length=256,
        verbose_name='Название региона',
    )
    analytical_table = models.FileField(
            upload_to=region_upload_path,
            max_length=1024,
            storage=OriginalNameStorage(),
            null=True,
            blank=True,
            verbose_name='Файл аналитических таблиц региона',
        )

    class Meta:
        verbose_name = "Регион"
        verbose_name_plural = "Регионы"

    def __str__(self):
        return self.name
