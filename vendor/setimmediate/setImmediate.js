"use strict";

(function attachSetImmediate(globalScope) {
  if (!globalScope) {
    return;
  }

  if (
    typeof globalScope.setImmediate === "function" &&
    typeof globalScope.clearImmediate === "function"
  ) {
    module.exports = {
      setImmediate: globalScope.setImmediate.bind(globalScope),
      clearImmediate: globalScope.clearImmediate.bind(globalScope),
    };
    return;
  }

  let nextHandle = 1;
  const tasks = new Map();
  let messageChannel = null;
  const messageQueue = [];

  const runTask = (handle) => {
    const task = tasks.get(handle);
    if (!task) {
      return;
    }

    tasks.delete(handle);
    task.callback(...task.args);
  };

  const installScheduler = () => {
    if (typeof queueMicrotask === "function") {
      return (handle) => queueMicrotask(() => runTask(handle));
    }

    if (typeof MessageChannel !== "undefined") {
      messageChannel = new MessageChannel();
      messageChannel.port1.onmessage = () => {
        const handle = messageQueue.shift();
        if (typeof handle === "number") {
          runTask(handle);
        }
      };

      return (handle) => {
        messageQueue.push(handle);
        messageChannel.port2.postMessage(0);
      };
    }

    return (handle) => setTimeout(() => runTask(handle), 0);
  };

  const schedule = installScheduler();

  function setImmediateShim(callback, ...args) {
    if (typeof callback !== "function") {
      throw new TypeError("setImmediate callback must be a function");
    }

    const handle = nextHandle++;
    tasks.set(handle, { callback, args });
    schedule(handle);
    return handle;
  }

  function clearImmediateShim(handle) {
    tasks.delete(handle);
  }

  globalScope.setImmediate = setImmediateShim;
  globalScope.clearImmediate = clearImmediateShim;

  module.exports = {
    setImmediate: setImmediateShim,
    clearImmediate: clearImmediateShim,
  };
})(
  typeof globalThis !== "undefined"
    ? globalThis
    : typeof self !== "undefined"
      ? self
      : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
          ? global
          : undefined,
);
