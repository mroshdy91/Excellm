"""Type definitions for inspection tools.

Stable output schemas using Pydantic for deterministic, versionable contracts.
"""

from typing import Any, Dict, List, Literal, Optional

from pydantic import BaseModel, Field

# Tool version - increment on breaking schema changes
TOOL_VERSION = "v1"

# Flag values as string literals for stability
FlagType = Literal[
    "EMPTY_OR_NEAR_EMPTY",
    "USED_RANGE_INFLATED",
    "EXTREME_USED_RANGE",
    "LIKELY_FORMAT_ONLY",
    "MULTI_REGION_SHEET",
    "OUTLIER_DATA_PRESENT",
    "DENSITY_INCONSISTENT",
    "HAS_TABLE_OBJECTS",
    "HAS_FILTER",
    "HAS_FREEZE_PANES",
    "MERGED_CELLS_PRESENT",
    "HAS_COMMENTS_NOTES",
    "HAS_FORMULAS",
    "HIDDEN_SHEET",
]

SheetStateType = Literal["visible", "hidden", "veryHidden"]
ModeType = Literal["quick", "deep"]
SeverityType = Literal["low", "medium", "high"]


# ============================================================================
# inspect_workbook Types
# ============================================================================


class LayoutFlags(BaseModel):
    """Sheet layout flags for quick inspection."""
    hasTableObjects: Optional[bool] = None
    hasAutoFilter: Optional[bool] = None
    hasFreezePanes: Optional[bool] = None
    mergedCellsCount: Optional[int] = None
    hiddenRowsCount: Optional[int] = None
    hiddenColsCount: Optional[int] = None
    commentsCount: Optional[int] = None
    formulasCount: Optional[int] = None


class SheetScore(BaseModel):
    """Score for sheet prioritization."""
    priority: float = Field(ge=0, le=1)
    density: Optional[float] = Field(default=None, ge=0, le=1)


class SheetIndex(BaseModel):
    """Per-sheet metadata in workbook inspection."""
    name: str
    state: SheetStateType
    protected: bool

    usedRangeReported: str
    realDataBounds: Optional[str] = None
    dataCellCount: Optional[int] = None
    nonEmptyRows: Optional[int] = None
    nonEmptyCols: Optional[int] = None

    layoutFlags: LayoutFlags
    flags: List[str] = Field(default_factory=list)
    score: SheetScore


class WorkbookInfo(BaseModel):
    """Workbook-level metadata."""
    name: str
    path: Optional[str] = None
    readOnly: bool
    protected: bool


class Recommendations(BaseModel):
    """Recommendations for next actions."""
    primaryCandidateSheets: List[str] = Field(default_factory=list)
    avoidSheets: List[str] = Field(default_factory=list)
    nextExploreSheet: str


class Meta(BaseModel):
    """Metadata for tool response."""
    tool: str
    version: str = TOOL_VERSION
    timestamp: str
    durationMs: int
    sheet: Optional[str] = None
    mode: Optional[ModeType] = None


class InspectWorkbookResult(BaseModel):
    """Complete response for inspect_workbook tool."""
    meta: Meta
    workbook: WorkbookInfo
    activeSheet: str
    sheetsIndex: List[SheetIndex]
    recommendations: Recommendations


# ============================================================================
# explore Types
# ============================================================================


class Region(BaseModel):
    """Detected data region."""
    id: str
    range: str
    density: Optional[float] = Field(default=None, ge=0, le=1)
    headerCandidateRows: List[int] = Field(default_factory=list)


class OutlierDistance(BaseModel):
    """Distance from primary region."""
    rows: int
    cols: int


class Outlier(BaseModel):
    """Detected outlier data block."""
    id: str
    range: str
    cellCount: Optional[int] = None
    distanceFromPrimary: Optional[OutlierDistance] = None


class DataFootprint(BaseModel):
    """Data footprint metrics."""
    usedRangeReported: str
    realDataBounds: Optional[str] = None
    dataCellCount: Optional[int] = None
    nonEmptyRows: Optional[int] = None
    nonEmptyCols: Optional[int] = None


class Layout(BaseModel):
    """Sheet layout information."""
    tables: Dict[str, Any]  # {"count": int|null, "names": [str]}
    mergedCellsCount: Optional[int] = None
    mergedRangesSample: List[str] = Field(default_factory=list)
    freezePanesAt: Optional[str] = None
    autoFilter: Optional[bool] = None
    hiddenRows: Optional[int] = None
    hiddenCols: Optional[int] = None


class ContentTypes(BaseModel):
    """Content type counts."""
    constants: Optional[int] = None
    formulas: Optional[int] = None


class CommentsNotes(BaseModel):
    """Comments/notes information."""
    count: Optional[int] = None
    rangesSample: List[str] = Field(default_factory=list)


class WriteSafety(BaseModel):
    """Write safety hints."""
    appendColumnToRegionId: Optional[str] = None
    avoidUsedRangeRightEdge: Optional[bool] = None


class ReadHints(BaseModel):
    """Suggestions for reading data."""
    primaryRegionId: Optional[str] = None
    suggestedHeaderScan: Optional[str] = None
    suggestedTailScan: Optional[str] = None
    suggestedBodyRead: Optional[str] = None
    suggestedOutlierScans: List[str] = Field(default_factory=list)
    suggestedRegionHeaderScans: Optional[Dict[str, str]] = None
    writeSafety: Optional[WriteSafety] = None


class RecommendationReason(BaseModel):
    """Reason for a recommendation."""
    flag: str
    severity: SeverityType
    why: str


class NextAction(BaseModel):
    """Suggested next action."""
    tool: str
    scope: Dict[str, str]
    mode: str


class ExploreRecommendations(BaseModel):
    """Recommendations from explore tool."""
    shouldRunDeep: bool
    reasons: List[RecommendationReason] = Field(default_factory=list)
    nextActions: List[NextAction] = Field(default_factory=list)


class ExploreResult(BaseModel):
    """Complete response for explore tool."""
    meta: Meta
    dataFootprint: DataFootprint
    regions: List[Region] = Field(default_factory=list)
    outliers: List[Outlier] = Field(default_factory=list)
    layout: Layout
    contentTypes: ContentTypes
    commentsNotes: CommentsNotes
    flags: List[str] = Field(default_factory=list)
    readHints: ReadHints
    recommendations: ExploreRecommendations


# ============================================================================
# Error Response
# ============================================================================


class ErrorResponse(BaseModel):
    """Structured error response."""
    success: Literal[False] = False
    error: Dict[str, Any]  # {"code": str, "message": str}
    meta: Meta
