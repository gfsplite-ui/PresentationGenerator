import os

# GigaChat Credentials
GIGACHAT_CREDENTIALS = os.environ.get(
    "GIGACHAT_CREDENTIALS",
    "MDE5ZGUzZmItZTM4My03Nzc3LTg3YzUtMmYyZTM5MzI4NzlhOmIxNDhlNjJjLWZkNDMtNDNhMC1hZTY0LTcxZDRkYmU2ODg3OA==",
)

# Путь к шаблону презентации (файл лежит рядом с кодом)
TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", "company_template.pptx")

# Индексы макетов
TITLE_LAYOUT_INDEX = 0
CONTENT_LAYOUT_INDEX = 13

# Добавлять ли изображения
ADD_IMAGES = True

# Радиус мягких краёв
SOFT_EDGE_RADIUS = 10
