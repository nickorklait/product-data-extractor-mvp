from typing import List, Optional

from pydantic import BaseModel, Field


class DocumentMetadata(BaseModel):
    source_file: Optional[str] = None
    product_name: Optional[str] = None
    supplier: Optional[str] = None
    plant: Optional[str] = None
    document_id: Optional[str] = None
    version: Optional[str] = None
    approval_status: Optional[str] = None
    approved_date: Optional[str] = None


class Nutrition(BaseModel):
    energy_kj: Optional[str] = None
    energy_kcal: Optional[str] = None
    fat_g: Optional[str] = None
    saturated_fat_g: Optional[str] = None
    carbohydrate_g: Optional[str] = None
    sugars_g: Optional[str] = None
    fibre_g: Optional[str] = None
    protein_g: Optional[str] = None
    salt_g: Optional[str] = None


class Allergens(BaseModel):
    allergens_contains: Optional[str] = None
    allergens_may_contain: Optional[str] = None
    allergen_statement: Optional[str] = None
    gluten: Optional[str] = None
    crustaceans: Optional[str] = None
    eggs: Optional[str] = None
    fish: Optional[str] = None
    peanuts: Optional[str] = None
    soybeans: Optional[str] = None
    milk: Optional[str] = None
    nuts: Optional[str] = None
    celery: Optional[str] = None
    mustard: Optional[str] = None
    sesame: Optional[str] = None
    sulphites: Optional[str] = None
    lupin: Optional[str] = None
    molluscs: Optional[str] = None


class ShelfLife(BaseModel):
    shelf_life_value: Optional[str] = None
    shelf_life_unit: Optional[str] = None
    shelf_life_text: Optional[str] = None


class Storage(BaseModel):
    storage_temperature: Optional[str] = None
    storage_text: Optional[str] = None


class PhysicalChemicalData(BaseModel):
    moisture: Optional[str] = None
    ph: Optional[str] = Field(default=None, alias="ph")
    density: Optional[str] = None
    particle_size: Optional[str] = None
    physical_chemical_text: Optional[str] = None


class KeyValueCandidate(BaseModel):
    section: Optional[str] = None
    key: str
    value: str
    source_line: Optional[int] = None
    source_reference: Optional[str] = None


class RawLine(BaseModel):
    line_number: int
    section: Optional[str] = None
    text: str
    source_reference: Optional[str] = None


class DetectedSection(BaseModel):
    section: str
    text: str


class DynamicField(BaseModel):
    field_name: str
    label: str
    section: Optional[str] = None
    value: Optional[str] = None
    unit: Optional[str] = None
    source: Optional[str] = None
    needs_review: bool = False


class SpecificationRow(BaseModel):
    section: str
    parameter: str
    value: Optional[str] = None
    min_value: Optional[str] = None
    max_value: Optional[str] = None
    unit: Optional[str] = None
    qualifier: Optional[str] = None
    source_line: Optional[int] = None
    source_reference: Optional[str] = None
    source_excerpt: Optional[str] = None
    needs_review: bool = False


class IPDAttributeRow(BaseModel):
    spec_id: Optional[str] = None
    node: str
    attribute: str
    attribute_description: Optional[str] = None
    data: Optional[str] = None
    uom: Optional[str] = None
    comments: Optional[str] = None
    source_reference: Optional[str] = None
    source_excerpt: Optional[str] = None
    status: str = "Not found"
    approved: bool = False


class ChangeLogEntry(BaseModel):
    changed_at: str
    actor: str = "Local reviewer"
    node: str
    attribute: str
    field: str
    old_value: Optional[str] = None
    new_value: Optional[str] = None


class ExtractionResult(BaseModel):
    metadata: DocumentMetadata
    nutrition: Nutrition
    allergens: Allergens
    shelf_life: ShelfLife
    storage: Storage
    physical_chemical: PhysicalChemicalData
    extraction_warnings: List[str] = []
    raw_text_excerpt: Optional[str] = None
    raw_text_full: Optional[str] = None
    raw_lines: List[RawLine] = []
    key_value_candidates: List[KeyValueCandidate] = []
    detected_sections: List[DetectedSection] = []
    review_flags: List[str] = []
    output_profile: str = "generic"
    dynamic_fields: List[DynamicField] = []
    specification_rows: List[SpecificationRow] = []
    ipd_rows: List[IPDAttributeRow] = []
    change_log: List[ChangeLogEntry] = []
