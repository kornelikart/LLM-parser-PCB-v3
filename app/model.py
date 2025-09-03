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
    
    foil_thickness: str = Field(
        default="",
        description="Толщина фольги"
    )
    
    layer_count: int = Field(
        default=0,
        description="Количество слоев печатной платы"
    )

    coverage_type: str = Field(
        default="",
        description="Финишное покрытие площадок"
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