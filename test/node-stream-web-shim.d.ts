// happy-dom's declarations reference `UnderlyingDefaultSource` from `node:stream/web`, which
// @types/node only exports from 25 on. This project holds @types/node on the Node 24 floor
// (see `.node-version` and renovate.json), so the name is re-exported here from the DOM
// global it mirrors. Delete this file when @types/node moves past 24.
declare module 'node:stream/web' {
	// oxlint-disable-next-line typescript/no-empty-object-type -- an augmentation can only merge an interface, not add a type alias.
	interface UnderlyingDefaultSource<R = unknown> extends globalThis.UnderlyingDefaultSource<R> {}
}

export {}
