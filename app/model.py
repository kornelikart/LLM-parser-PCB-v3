from pydantic import BaseModel, Field

class PCBCharacteristics(BaseModel):
    company_name: str = Field(
        default="",
        description="Название компании производителя печатной платы"
    )
    
    board_name: str = Field(
        default="",
        description="Название печатной платы"
    )

    quantity: str = Field(
        default="",
        description=(
            "Количество плат в заказе (Qty / 'Количество плат в заказе, шт' / "
            "'Требуется изготовить, шт' / 'Количество, шт'). Указывать как в документе, "
            "например '3000 pcs (1500 panels)', 'Кратно 4'."
        )
    )

    base_material: str = Field(
        default="",
        description="Материал основания печатной платы"
    )

    board_thickness: str = Field(
        default="",
        description="Толщина печатной платы"
    )

    foil_thickness: str = Field(
        default="",
        description=(
            "Толщина МЕДНОЙ ФОЛЬГИ внешних слоёв (CU foil / Copper thickness / Толщина меди / Thickness CU). "
            "Значение может быть в мм (0.018 мм = 18 мкм = 0.5 oz) или в мкм/OZ. "
            "ВАЖНО: НЕ путать с толщиной маски (Solder mask / Маска) — маску игнорировать. "
            "Примеры: '0.018 мм', '35 мкм', '1 OZ (35 мкм)', '0.5 oz', '18 um + 17 um plating'."
        )
    )
    
    layer_count: int = Field(
        default=0,
        description="Количество слоев печатной платы"
    )

    coverage_type: str = Field(
        default="",
        description=(
            "Финишное покрытие поверхности платы (Surface Finish / Финишное покрытие). "
            "Примеры: ENIG, HASL, HASL LF, OSP, Imm. gold, Imm. silver, Imm. tin, "
            "Hard gold, Soft gold, ENEPIG, Иммерсионное золото, Химическое олово, ОСП. "
            "Указывать точно как в документе."
        )
    )
    
    board_size: str = Field(
        default="",
        description="Размер печатной платы"
    )
    
    panelization: str = Field(
        default="",
        description="Панелизация печатной платы (общее описание)"
    )

    panel_size: str = Field(
        default="",
        description=(
            "ТОЛЬКО габариты панели/заготовки: длина x ширина в мм "
            "('Размер панели, мм', 'Габариты панели', 'Panel size'). "
            "Формат: 'ДЛИНА x ШИРИНА', например '134 x 74'. "
            "НЕ путать с размером одиночной платы (board_size). "
            "Если панелизации нет — пустая строка."
        )
    )

    boards_per_panel: str = Field(
        default="",
        description=(
            "Количество печатных плат в одной панели/заготовке "
            "('Количество ПП в панели', 'Количество плат в панели', 'Boards per panel'). "
            "Только число, например '1', '2', '4'. Если панель из матрицы NxM — "
            "укажите произведение (2x3 → '6')."
        )
    )

    different_boards_per_panel: str = Field(
        default="",
        description=(
            "Количество РАЗНЫХ типов плат на панели (мультипанель/сборная панель). "
            "Если на панели повторяется одна и та же плата — '1'. "
            "Если панель содержит разные платы (напр. 2 типа) — их количество."
        )
    )

    technological_fields: str = Field(
        default="",
        description=(
            "Наличие технологических полей (полей по краю панели, рамки, "
            "'Наличие технологических полей', 'Размер полей', 'тех. поля', "
            "'ширина полей'): Yes/No. Если панель больше платы за счёт полей — Yes."
        )
    )

    impedance_control: str = Field(
        default="",
        description=(
            "Контроль волнового сопротивления (импеданса): "
            "'Контроль импедансов', 'Контроль волнового сопротивления', "
            "'impedance control', 'controlled impedance'. "
            "Значения: 'Yes'/'No'. 'есть'/'требуется' → Yes, 'нет' → No. "
            "Если в документе есть таблица импедансов со значениями (Ом) — Yes."
        )
    )

    min_hole_size: str = Field(
        default="",
        description=(
            "Минимальный диаметр металлизированного отверстия, мм: "
            "'Мин. диаметр металлизир. отверстия', 'Минимальное металлизированное отверстие', "
            "'Smallest plated hole size', 'Минимальный диаметр сквозного металлизированного отверстия'. "
            "Только число в мм, например '0.2', '0.3'."
        )
    )

    marking_side: str = Field(
        default="",
        description=(
            "Сторона нанесения маркировки изготовителя: "
            "'сторона маркировки', 'Маркировка изготовителя', 'Маркировка'. "
            "Значения: 'TOP', 'BOTTOM', 'TOP+BOTTOM', 'None'. "
            "Русские варианты: 'сверху' → TOP, 'снизу' → BOTTOM, "
            "'с двух сторон' → TOP+BOTTOM, 'отсутствует' → None."
        )
    )

    serial_number: str = Field(
        default="",
        description=(
            "Требование индивидуального серийного/порядкового номера на каждой плате: "
            "'порядковый номер в партии', 'заводской порядковый номер', "
            "'уникальный код', 'серийный номер', 'serial number', 'barcode', "
            "'лазерный штрих-код'. Значения: 'Yes'/'No'. "
            "Если требуется формат номера (напр. 'требуется в формате ХХХ') — Yes."
        )
    )

    solder_mask_colour: str = Field(
        default="",
        description="Наличие маски /цвет"
    )

    solder_mark_colour: str = Field(
        default="",
        description="Наличие маркировки маркировочной краской/цвет"
    )

    soldering_surface: str = Field(
        default="",
        description="Монтаж печатных плат"
    )

    electrical_testing: str = Field(
        default="",
        description="Электротестирование"
    )

    edge_plating: str = Field(
        default="",
        description="Металлизированный торец платы"
    )

    contour_treatment: str = Field(
        default="",
        description="Мех обработка контура"
    )

    pcb_type: str = Field(
        default="",
        description="Тип печатной платы (Rigid, Flex, Rigid-Flex, Semi-Flex и т.д.)"
    )

    peelable_mask: str = Field(
        default="",
        description="Пилинг-маска (Yes/No)"
    )

    gold_fingers: str = Field(
        default="",
        description="Золотые контактные пальцы (Yes/No)"
    )

    ipc_class: str = Field(
        default="",
        description="Класс качества IPC (IPC Class 2 или IPC Class 3)"
    )

    back_drill: str = Field(
        default="",
        description="Обратное сверление back drill (Yes/No)"
    )

    flex_type: str = Field(
        default="",
        description="Тип гибкой платы: None, Single Side, double side, Multilayer"
    )

    cover_layer: str = Field(
        default="",
        description="Покрывающий слой Coverlay (Yes/No)"
    )

    flex_layer_location: str = Field(
        default="",
        description="Расположение гибких слоёв: None, Inner Layer, Outer Layer"
    )

    coin: str = Field(
        default="",
        description="Встроенные медные вставки Coin (Yes/No)"
    )

    embedded_components: str = Field(
        default="",
        description="Встроенные компоненты Embedded Components (Yes/No)"
    )