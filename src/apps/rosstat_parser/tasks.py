import sys
import traceback
from datetime import datetime

from celery import shared_task
from django.core.files import File
from apps.rosstat_parser.models import Region
from apps.rosstat_parser.api.v1.services.excel_writer import create_all_tables
from apps.rosstat_parser.api.v1.services.downloader import download_rosstat_tables


def log_progress(message: str) -> None:
    """Пишет прогресс напрямую в stderr контейнера, минуя все перехваты."""
    stream = sys.__stderr__
    stream.write(f"[{datetime.now():%H:%M:%S}] {message}\n")
    stream.flush()


REGIONS = [
        "Республика Адыгея",
        "Республика Алтай",
        "Республика Башкортостан",
        "Республика Бурятия",
        "Республика Дагестан",
        "Республика Ингушетия",
        "Кабардино-Балкарская Республика",
        "Республика Калмыкия",
        "Карачаево-Черкесская Республика",
        "Республика Карелия",
        "Республика Коми",
        "Республика Крым",
        "Республика Марий Эл",
        "Республика Мордовия",
        "Республика Саха (Якутия)",
        "Республика Северная Осетия - Алания",
        "Республика Татарстан",
        "Республика Тыва",
        "Удмуртская Республика",
        "Республика Хакасия",
        "Чеченская Республика",
        "Чувашская Республика",
        "Алтайский край",
        "Забайкальский край",
        "Камчатский край",
        "Краснодарский край",
        "Красноярский край",
        "Пермский край",
        "Приморский край",
        "Ставропольский край",
        "Хабаровский край",
        "Амурская область",
        "Архангельская область",
        "Астраханская область",
        "Белгородская область",
        "Брянская область",
        "Владимирская область",
        "Волгоградская область",
        "Вологодская область",
        "Воронежская область",
        "Ивановская область",
        "Иркутская область",
        "Калининградская область",
        "Калужская область",
        "Кемеровская область - Кузбасс",
        "Кировская область",
        "Костромская область",
        "Курганская область",
        "Курская область",
        "Ленинградская область",
        "Липецкая область",
        "Магаданская область",
        "Московская область",
        "Мурманская область",
        "Нижегородская область",
        "Новгородская область",
        "Новосибирская область",
        "Омская область",
        "Оренбургская область",
        "Орловская область",
        "Пензенская область",
        "Псковская область",
        "Ростовская область",
        "Рязанская область",
        "Самарская область",
        "Саратовская область",
        "Сахалинская область",
        "Свердловская область",
        "Смоленская область",
        "Тамбовская область",
        "Тверская область",
        "Томская область",
        "Тульская область",
        "Тюменская область",
        "Ульяновская область",
        "Челябинская область",
        "Ярославская область",
        "г.Москва",
        "г.Санкт-Петербург",
        "г.Севастополь",
        "Еврейская автономная область",
        "Ненецкий автономный округ",
        "Ханты-Мансийский автономный округ - Югра",
        "Чукотский автономный округ",
        "Ямало-Ненецкий автономный округ",
        "Тюменская область (кроме Ханты-Мансийского автономного округа-Югры и Ямало-Ненецкого автономного округа)",
    ]

@shared_task
def create_region_table():
    log_progress(f"=== Старт задачи create_region_table, регионов: {len(REGIONS)} ===")
    Region.objects.all().delete()
    log_progress("Скачиваю файлы Росстата...")
    download_rosstat_tables(save_dir="downloads")
    log_progress("Файлы скачаны, начинаю формировать таблицы.")

    failed_regions = []
    succeeded_regions = 0

    for idx, region in enumerate(REGIONS, start=1):
        log_progress(f"[{idx}/{len(REGIONS)}] Обработка региона: {region}")
        try:
            create_all_tables(region)

            file_path = f"tables/Аналитические таблицы {region}.xlsx"
            region_obj = Region.objects.create(name=region)

            with open(file_path, "rb") as f:
                region_obj.analytical_table.save(
                    f"Аналитические таблицы {region}.xlsx",
                    File(f)
                )
                region_obj.save()
        except Exception as exc:
            failed_regions.append((region, repr(exc)))
            log_progress(
                f"[{idx}/{len(REGIONS)}] ОШИБКА на регионе {region}: {exc!r} "
                f"(пропускаем и идём дальше)"
            )
            traceback.print_exc(file=sys.__stderr__)
            sys.__stderr__.flush()
            continue

        succeeded_regions += 1
        log_progress(f"[{idx}/{len(REGIONS)}] Регион сохранён: {region}")

    log_progress(
        f"=== Задача create_region_table завершена. "
        f"Успешно: {succeeded_regions}/{len(REGIONS)}, "
        f"с ошибками: {len(failed_regions)} ==="
    )
    if failed_regions:
        log_progress("Регионы с ошибками:")
        for region, err in failed_regions:
            log_progress(f"  - {region}: {err}")

    return {
        "total": len(REGIONS),
        "succeeded": succeeded_regions,
        "failed": [r for r, _ in failed_regions],
    }
