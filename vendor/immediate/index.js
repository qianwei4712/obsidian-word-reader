"use strict";

let messageChannelScheduler = null;

function createMessageChannelScheduler() {
  if (typeof MessageChannel === "undefined") {
    return null;
  }

  const queue = [];
  const channel = new MessageChannel();
  channel.port1.onmessage = () => {
    const task = queue.shift();
    if (!task) {
      return;
    }

    task();
  };

  return (task) => {
    queue.push(task);
    channel.port2.postMessage(0);
  };
}

function scheduleTask(task) {
  if (typeof queueMicrotask === "function") {
    queueMicrotask(task);
    return;
  }

  if (typeof Promise !== "undefined") {
    Promise.resolve()
      .then(task)
      .catch((error) => {
        setTimeout(() => {
          throw error;
        }, 0);
      });
    return;
  }

  if (!messageChannelScheduler) {
    messageChannelScheduler = createMessageChannelScheduler();
  }

  if (messageChannelScheduler) {
    messageChannelScheduler(task);
    return;
  }

  setTimeout(task, 0);
}

module.exports = function immediate(task) {
  if (typeof task !== "function") {
    throw new TypeError("immediate task must be a function");
  }

  scheduleTask(task);
};
