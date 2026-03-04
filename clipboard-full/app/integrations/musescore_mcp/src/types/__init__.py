"""Type definitions for MuseScore MCP."""

from .action_types import *

__all__ = [
    "ActionSequence",
    "getScoreAction",
    "addNoteAction", 
    "addRestAction",
    "addTupletAction",
    "addLyricsAction",
    "addInstrumentAction",
    "setStaffMuteAction",
    "setInstrumentSoundAction",
    "appendMeasureAction",
    "deleteSelectionAction",
    "getCursorInfoAction",
    "goToMeasureAction",
    "nextElementAction",
    "prevElementAction",
    "selectCurrentMeasureAction",
    "insertMeasureAction",
    "goToFinalMeasureAction",
    "goToBeginningOfScoreAction",
    "setTimeSignatureAction",
    "undoAction",
    "nextStaffAction",
    "prevStaffAction"
]