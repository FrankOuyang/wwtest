from pydantic import BaseModel
from typing import List, Optional

class Purpose(BaseModel):
    title: str
    content: str

class Nature(BaseModel):
    title: str
    content: str

class Pathogen(BaseModel):
    title: str
    content: str

class TestArticles(BaseModel):
    title: str
    items: List[str]

class Control(BaseModel):
    title: str
    content: str

class ExperimentalSys(BaseModel):
    title: str
    quantity: str
    species: str
    male: str = ""
    female: str

class MethodSection(BaseModel):
    title: str
    steps: List[str]
    content: str

class Methods(BaseModel):
    title: str
    sections: List[MethodSection]

class ExperimentTimeline(BaseModel):
    title: str
    image: str

class ExperimentalGroups(BaseModel):
    title: str
    columns: List[str]
    rows: List[List[str]]

class DetectionIndicatorItem(BaseModel):
    title: str
    content: str

class DetectionIndicators(BaseModel):
    title: str
    items: List[DetectionIndicatorItem]

class ExperimentProtocol(BaseModel):
    purpose: Purpose
    nature: Nature
    pathogen: Pathogen
    test_articles: TestArticles
    positive_control: Control
    solvent_control: Control
    experimental_sys: ExperimentalSys
    methods: Methods
    experiment_timeline: ExperimentTimeline
    experimental_groups: ExperimentalGroups
    detection_indicators: DetectionIndicators