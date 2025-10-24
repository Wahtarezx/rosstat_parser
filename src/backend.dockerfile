FROM python:3.11

ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1
ENV DJANGO_SETTINGS_MODULE 'config.settings'

WORKDIR /opt/src

COPY pyproject.toml pyproject.toml
RUN mkdir -p /opt/src/static/ && \
    mkdir -p /opt/src/media/  &&  \
    pip install --upgrade pip && \
    pip install 'poetry>=1.4.2' && \
    poetry config virtualenvs.create false && \
    poetry install --no-root --only main

RUN apt-get update && apt-get install -y \
    wget curl unzip gnupg ca-certificates \
    fonts-liberation libasound2 libatk-bridge2.0-0 libatk1.0-0 libcups2 \
    libdbus-1-3 libgdk-pixbuf-2.0-0 libnspr4 libnss3 libx11-6 \
    libxcomposite1 libxdamage1 libxext6 libxfixes3 libxrandr2 xdg-utils \
    chromium chromium-driver \
    && rm -rf /var/lib/apt/lists/*

COPY . .

EXPOSE 8000

COPY ./entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]
