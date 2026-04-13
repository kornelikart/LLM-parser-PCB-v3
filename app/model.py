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
        description="Панелизация печатной платы"
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