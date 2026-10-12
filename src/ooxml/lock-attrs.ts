/**
 * The attributes each DrawingML locking element defines, in emit order (ECMA-376 §20.1.2.2.x /
 * §20.1.2.2.34): `a:spLocks`, `a:picLocks`, `a:graphicFrameLocks` and `a:grpSpLocks`.
 *
 * Here rather than in `gen/drawingml/locks.ts` because the read side's `readModelToIr` needs the
 * same list to tell an `objectLock` flag it can write from one it cannot. No runtime imports.
 */
export const SHAPE_LOCK_ATTRS = [
	'noGrp',
	'noSelect',
	'noRot',
	'noChangeAspect',
	'noMove',
	'noResize',
	'noEditPoints',
	'noAdjustHandles',
	'noChangeArrowheads',
	'noChangeShapeType',
	'noTextEdit',
] as const
export const PICTURE_LOCK_ATTRS = [
	'noGrp',
	'noSelect',
	'noRot',
	'noChangeAspect',
	'noMove',
	'noResize',
	'noEditPoints',
	'noAdjustHandles',
	'noChangeArrowheads',
	'noChangeShapeType',
	'noCrop',
] as const
export const GRAPHIC_FRAME_LOCK_ATTRS = [
	'noGrp',
	'noDrilldown',
	'noSelect',
	'noChangeAspect',
	'noMove',
	'noResize',
] as const
export const GROUP_SHAPE_LOCK_ATTRS = [
	'noGrp',
	'noUngrp',
	'noSelect',
	'noRot',
	'noChangeAspect',
	'noMove',
	'noResize',
] as const

/** Every attribute name the four tables between them list. */
