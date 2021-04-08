#  -*- coding: UTF-8 -*-
# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2006-2021 NV Access Limited, Peter Vágner, Aleksey Sadovoy, Babbage B.V., Bill Dengler,
# Julien Cochuyt

import re
import weakref
import garbageHandler
from logHandler import log
import config
import controlTypes
import api
import textInfos
import queueHandler
import winKernel

from .commands import (
	# Commands that are used in this file.
	LangChangeCommand,
	EndUtteranceCommand,
	CallbackCommand,
)

from .types import (
	SpeechSequence,
	logBadSequenceTypes,
	GeneratorWithReturn,
	_flattenNestedSequences,
)

from typing import (
	Optional,
	Generator,
	Callable,
	Literal,
	TYPE_CHECKING,
)

from .speech import (
	getTextInfoSpeech,
	speak,
	SpeakTextInfoState,
	speakObject,
)

if TYPE_CHECKING:
	import NVDAObjects

CURSOR_CARET = 0
CURSOR_REVIEW = 1
CURSOR = Literal[CURSOR_REVIEW, CURSOR_CARET]
SayAllHandler = None


def initalize():
	global SayAllHandler
	SayAllHandler = _SayAllHandler()


def _yieldIfNonEmpty(seq: SpeechSequence):
	"""Helper method to yield the sequence if it is not None or empty."""
	if seq:
		yield seq


class SpeechWithoutPauses:
	_pendingSpeechSequence: SpeechSequence
	re_last_pause = re.compile(
		r"^(.*(?<=[^\s.!?])[.!?][\"'”’)]?(?:\s+|$))(.*$)",
		re.DOTALL | re.UNICODE
	)

	def __init__(
			self,
			speakFunc: Callable[[SpeechSequence], None]
	):
		"""
		:param speakFunc: Function used by L{speakWithoutPauses} to speak. This will likely be speech.speak.
		"""
		self.speak = speakFunc
		self.reset()

	def reset(self):
		self._pendingSpeechSequence = []

	def speakWithoutPauses(
			self,
			speechSequence: Optional[SpeechSequence],
			detectBreaks: bool = True
	) -> bool:
		"""
		Speaks the speech sequences given over multiple calls,
		only sending to the synth at acceptable phrase or sentence boundaries,
		or when given None for the speech sequence.
		@return: C{True} if something was actually spoken,
			C{False} if only buffering occurred.
		"""
		speech = GeneratorWithReturn(self.getSpeechWithoutPauses(
			speechSequence,
			detectBreaks
		))
		for seq in speech:
			self.speak(seq)
		return speech.returnValue

	def getSpeechWithoutPauses(  # noqa: C901
			self,
			speechSequence: Optional[SpeechSequence],
			detectBreaks: bool = True
	) -> Generator[SpeechSequence, None, bool]:
		"""
		Generate speech sequences over multiple calls,
		only returning a speech sequence at acceptable phrase or sentence boundaries,
		or when given None for the speech sequence.
		@return: The speech sequence that can be spoken without pauses. The 'return' for this generator function,
		is a bool which indicates whether this sequence should be considered valid speech. Use
		L{GeneratorWithReturn} to retain the return value. A generator is used because the previous
		implementation had several calls to speech, this approach replicates that.
		"""
		if speechSequence is not None:
			logBadSequenceTypes(speechSequence)
		# Break on all explicit break commands
		if detectBreaks and speechSequence:
			speech = GeneratorWithReturn(self._detectBreaksAndGetSpeech(speechSequence))
			yield from speech
			return speech.returnValue  # Don't fall through to flush / normal speech

		if speechSequence is None:  # Requesting flush
			pending = self._flushPendingSpeech()
			yield from _yieldIfNonEmpty(pending)
			return bool(pending)  # Don't fall through to handle normal speech

		# Handling normal speech
		speech = self._getSpeech(speechSequence)
		yield from _yieldIfNonEmpty(speech)
		return bool(speech)

	def _detectBreaksAndGetSpeech(
			self,
			speechSequence: SpeechSequence
	) -> Generator[SpeechSequence, None, bool]:
		lastStartIndex = 0
		sequenceLen = len(speechSequence)
		gotValidSpeech = False
		for index, item in enumerate(speechSequence):
			if isinstance(item, EndUtteranceCommand):
				if index > 0 and lastStartIndex < index:
					subSequence = speechSequence[lastStartIndex:index]
					yield from _yieldIfNonEmpty(
						self._getSpeech(subSequence)
					)
				yield from _yieldIfNonEmpty(
					self._flushPendingSpeech()
				)
				gotValidSpeech = True
				lastStartIndex = index + 1
		if lastStartIndex < sequenceLen:
			subSequence = speechSequence[lastStartIndex:]
			seq = self._getSpeech(subSequence)
			gotValidSpeech = bool(seq)
			yield from _yieldIfNonEmpty(seq)
		return gotValidSpeech

	def _flushPendingSpeech(self) -> SpeechSequence:
		"""
		@return: may be empty sequence
		"""
		# Place the last incomplete phrase in to finalSpeechSequence to be spoken now
		pending = self._pendingSpeechSequence
		self._pendingSpeechSequence = []
		return pending

	def _getSpeech(
			self,
			speechSequence: SpeechSequence
	) -> SpeechSequence:
		"""
		@return: May be an empty sequence
		"""
		finalSpeechSequence: SpeechSequence = []  # To be spoken now
		pendingSpeechSequence: speechSequence = []  # To be saved off for speaking later
		# Scan the given speech and place all completed phrases in finalSpeechSequence to be spoken,
		# And place the final incomplete phrase in pendingSpeechSequence
		for index in range(len(speechSequence) - 1, -1, -1):
			item = speechSequence[index]
			if isinstance(item, str):
				m = self.re_last_pause.match(item)
				if m:
					before, after = m.groups()
					if after:
						pendingSpeechSequence.append(after)
					if before:
						finalSpeechSequence.extend(self._flushPendingSpeech())
						finalSpeechSequence.extend(speechSequence[0:index])
						finalSpeechSequence.append(before)
						# Apply the last language change to the pending sequence.
						# This will need to be done for any other speech change commands introduced in future.
						for changeIndex in range(index - 1, -1, -1):
							change = speechSequence[changeIndex]
							if not isinstance(change, LangChangeCommand):
								continue
							pendingSpeechSequence.append(change)
							break
						break
				else:
					pendingSpeechSequence.append(item)
			else:
				pendingSpeechSequence.append(item)
		if pendingSpeechSequence:
			pendingSpeechSequence.reverse()
			self._pendingSpeechSequence.extend(pendingSpeechSequence)
		return finalSpeechSequence


class _SayAllHandler:
	def __init__(self):
		self.lastSayAllMode = None
		self.speechWithoutPausesInstance = SpeechWithoutPauses(speakFunc=speak)
		#: The active say all manager.
		#: This is a weakref because the manager should be allowed to die once say all is complete.
		self._getActiveSayAll = lambda: None  # noqa: Return None when called like a dead weakref.

	def stop(self):
		'''
		Stops any active objects reader and resets the SayAllHandler's speech.SpeechWithoutPauses instance
		'''
		active = self._getActiveSayAll()
		if active:
			active.stop()
		self.speechWithoutPausesInstance.reset()

	def isRunning(self):
		"""Determine whether say all is currently running.
		@return: C{True} if say all is currently running, C{False} if not.
		@rtype: bool
		"""
		return bool(self._getActiveSayAll())

	def readObjects(self, obj: 'NVDAObjects.NVDAObject'):
		reader = _ObjectsReader(self, obj)
		self._getActiveSayAll = weakref.ref(reader)
		reader.next()

	def readText(self, cursor: CURSOR):
		self.lastSayAllMode = cursor
		try:
			reader = _TextReader(self, cursor)
		except NotImplementedError:
			log.debugWarning("Unable to make reader", exc_info=True)
			return
		self._getActiveSayAll = weakref.ref(reader)
		reader.nextLine()


class _ObjectsReader(garbageHandler.TrackedObject):

	def __init__(self, handler: _SayAllHandler, root: 'NVDAObjects.NVDAObject'):
		self.handler = handler
		self.walker = self.walk(root)
		self.prevObj = None

	def walk(self, obj: 'NVDAObjects.NVDAObject'):
		yield obj
		child = obj.simpleFirstChild
		while child:
			for descendant in self.walk(child):
				yield descendant
			child = child.simpleNext

	def next(self):
		if not self.walker:
			# We were stopped.
			return
		if self.prevObj:
			# We just started speaking this object, so move the navigator to it.
			api.setNavigatorObject(self.prevObj, isFocus=self.handler.lastSayAllMode == CURSOR_CARET)
			winKernel.SetThreadExecutionState(winKernel.ES_SYSTEM_REQUIRED)
		# Move onto the next object.
		self.prevObj = obj = next(self.walker, None)
		if not obj:
			return
		# Call this method again when we start speaking this object.
		callbackCommand = CallbackCommand(self.next, name="say-all:next")
		speakObject(obj, reason=controlTypes.OutputReason.SAYALL, _prefixSpeechCommand=callbackCommand)

	def stop(self):
		self.walker = None


class _TextReader(garbageHandler.TrackedObject):
	"""Manages continuous reading of text.
	This is intended for internal use only.

	The high level flow of control is as follows:
	1. The constructor sets things up.
	2. L{nextLine} is called to read the first line.
	3. When it speaks a line, L{nextLine} request that L{lineReached} be called
		when we start speaking this line, providing the position and state at this point.
	4. When we start speaking a line, L{lineReached} is called
		and moves the cursor to that line.
	5. L{lineReached} calls L{nextLine}.
	6. If there are more lines, L{nextLine} works as per steps 3 and 4.
	7. Otherwise, if the object doesn't support page turns, we're finished.
	8. If the object does support page turns,
		we request that L{turnPage} be called when speech is finished.
	9. L{turnPage} tries to turn the page.
	10. If there are no more pages, we're finished.
	11. If there is another page, L{turnPage} calls L{nextLine}.
	"""
	MAX_BUFFERED_LINES = 10

	def __init__(self, handler: _SayAllHandler, cursor: CURSOR):
		self.handler = handler
		self.cursor = cursor
		self.trigger = SayAllProfileTrigger()
		self.reader = None
		# Start at the cursor.
		if cursor == CURSOR_CARET:
			try:
				self.reader = api.getCaretObject().makeTextInfo(textInfos.POSITION_CARET)
			except (NotImplementedError, RuntimeError) as e:
				raise NotImplementedError("Unable to make TextInfo: " + str(e))
		else:
			self.reader = api.getReviewPosition()
		# #10899: SayAll profile can't be activated earlier because they may not be anything to read
		self.trigger.enter()
		self.speakTextInfoState = SpeakTextInfoState(self.reader.obj)
		self.numBufferedLines = 0

	def nextLine(self):
		if not self.reader:
			log.debug("no self.reader")
			# We were stopped.
			return
		if not self.reader.obj:
			log.debug("no self.reader.obj")
			# The object died, so we should too.
			self.finish()
			return
		bookmark = self.reader.bookmark
		# Expand to the current line.
		# We use move end rather than expand
		# because the user might start in the middle of a line
		# and we don't want to read from the start of the line in that case.
		# For lines after the first, it's also more efficient because
		# we're already at the start of the line, so there's no need to search backwards.
		delta = self.reader.move(textInfos.UNIT_READINGCHUNK, 1, endPoint="end")
		if delta <= 0:
			# No more text.
			if isinstance(self.reader.obj, textInfos.DocumentWithPageTurns):
				# Once the last line finishes reading, try turning the page.
				cb = CallbackCommand(self.turnPage, name="say-all:turnPage")
				self.handler.speechWithoutPausesInstance.speakWithoutPauses([cb, EndUtteranceCommand()])
			else:
				self.finish()
			return

		# Copy the speakTextInfoState so that speak callbackCommand
		# and its associated callback are using a copy isolated to this specific line.
		state = self.speakTextInfoState.copy()
		# Call lineReached when we start speaking this line.
		# lineReached will move the cursor and trigger reading of the next line.

		def _onLineReached(obj=self.reader.obj, state=state):
			self.lineReached(obj, bookmark, state)

		cb = CallbackCommand(
			_onLineReached,
			name="say-all:lineReached"
		)

		# Generate the speech sequence for the reader textInfo
		# and insert the lineReached callback at the very beginning of the sequence.
		# _linePrefix on speakTextInfo cannot be used here
		# As it would be inserted in the sequence after all initial control starts which is too late.
		speechGen = getTextInfoSpeech(
			self.reader,
			unit=textInfos.UNIT_READINGCHUNK,
			reason=controlTypes.OutputReason.SAYALL,
			useCache=state
		)
		seq = list(_flattenNestedSequences(speechGen))
		seq.insert(0, cb)
		# Speak the speech sequence.
		spoke = self.handler.speechWithoutPausesInstance.speakWithoutPauses(seq)
		# Update the textInfo state ready for when speaking the next line.
		self.speakTextInfoState = state.copy()

		# Collapse to the end of this line, ready to read the next.
		try:
			self.reader.collapse(end=True)
		except RuntimeError:
			# This occurs in Microsoft Word when the range covers the end of the document.
			# without this exception to indicate that further collapsing is not possible,
			# say all could enter an infinite loop.
			self.finish()
			return
		if not spoke:
			# This line didn't include a natural pause, so nothing was spoken.
			self.numBufferedLines += 1
			if self.numBufferedLines < self.MAX_BUFFERED_LINES:
				# Move on to the next line.
				# We queue this to allow the user a chance to stop say all.
				queueHandler.queueFunction(queueHandler.eventQueue, self.nextLine)
			else:
				# We don't want to buffer too much.
				# Force speech. lineReached will resume things when speech catches up.
				self.handler.speechWithoutPausesInstance.speakWithoutPauses(None)
				# The first buffered line has now started speaking.
				self.numBufferedLines -= 1

	def lineReached(self, obj, bookmark, state):
		# We've just started speaking this line, so move the cursor there.
		state.updateObj()
		updater = obj.makeTextInfo(bookmark)
		if self.cursor == CURSOR_CARET:
			updater.updateCaret()
		if self.cursor != CURSOR_CARET or config.conf["reviewCursor"]["followCaret"]:
			api.setReviewPosition(updater, isCaret=self.cursor == CURSOR_CARET)
		winKernel.SetThreadExecutionState(winKernel.ES_SYSTEM_REQUIRED)
		if self.numBufferedLines == 0:
			# This was the last line spoken, so move on.
			self.nextLine()
		else:
			self.numBufferedLines -= 1

	def turnPage(self):
		try:
			self.reader.obj.turnPage()
		except RuntimeError:
			log.debug("No more pages")
			# No more pages.
			self.stop()
			return
		self.reader = self.reader.obj.makeTextInfo(textInfos.POSITION_FIRST)
		self.nextLine()

	def finish(self):
		# There is no more text.
		# Call stop to clean up, but only after speech completely finishes.
		# Otherwise, if a different synth is being used for say all,
		# we might switch synths too early and truncate the final speech.
		# We do this by putting a CallbackCommand at the start of a new utterance.
		cb = CallbackCommand(self.stop, name="say-all:stop")
		self.handler.speechWithoutPausesInstance.speakWithoutPauses([
			EndUtteranceCommand(),
			cb,
			EndUtteranceCommand()
		])

	def stop(self):
		if not self.reader:
			return
		self.reader = None
		self.trigger.exit()
		self.trigger = None

	def __del__(self):
		self.stop()


class SayAllProfileTrigger(config.ProfileTrigger):
	"""A configuration profile trigger for when say all is in progress.
	"""
	spec = "sayAll"
