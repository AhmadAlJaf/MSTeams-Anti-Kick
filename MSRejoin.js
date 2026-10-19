// ==UserScript==
// @name         Microsoft Teams Ultra-Fast Auto Rejoin MASTERED
// @namespace    https://teams.microsoft.com/
// @version      3.1.0
// @description  Automatic rejoin for personal and work/school Teams meetings.
// @match        https://teams.microsoft.com/*
// @match        https://teams.cloud.microsoft/*
// @match        https://teams.live.com/*
// @run-at       document-idle
// @grant        none
// @author       AhmadAlJaf
// @noframes
// ==/UserScript==

(() => {
  'use strict';

  if (!['teams.microsoft.com', 'teams.cloud.microsoft', 'teams.live.com']
    .includes(location.hostname)) return;

  // Stop an earlier helper if this version is run again in the same page.
  window.__teamsAutoRejoin?.stop?.();

  // All times are milliseconds. Retries do not delay the first click.
  const CONFIG = {
    pollInterval: 100,
    rejoinRetryDelay: 1000,
    mediaRetryDelay: 750,
    joinRetryDelay: 3000,
    pendingTimeout: 120000,
    // Also finish a prejoin screen when Teams is opened or reloaded.
    joinOnOpen: true
  };

  const normalize = value => String(value || '')
    .replace(/[’‘]/g, "'").replace(/\s+/g, ' ').trim();

  const label = el => normalize(
    el.getAttribute('aria-label') ||
    (el.getAttribute('aria-labelledby') || '').split(/\s+/)
      .map(id => document.getElementById(id)?.textContent || '').join(' ').trim() ||
    el.innerText || el.textContent
  );

  const visible = el => Boolean(
    el?.isConnected && el.getClientRects().length &&
    !el.closest('[hidden], [inert], [aria-hidden="true"]') &&
    getComputedStyle(el).display !== 'none' &&
    getComputedStyle(el).visibility !== 'hidden'
  );

  const enabled = el => Boolean(el) && !el.disabled &&
    el.getAttribute('aria-disabled') !== 'true';

  const checked = el => {
    if (!el) return null;
    if (el instanceof HTMLInputElement &&
      /^(checkbox|radio)$/.test(el.type)) return el.checked;
    const value = el.getAttribute('aria-checked') ??
      el.getAttribute('aria-pressed');
    return value === 'true' ? true : value === 'false' ? false : null;
  };

  const unique = elements => elements.length === 1 ? elements[0] : null;

  function mediaState(type) {
    const controls = [...document.querySelectorAll(
      'button, [role="button"], [role="switch"], input[type="checkbox"]'
    )].filter(visible);

    const onPattern = type === 'camera'
      ? /^turn camera off\b/i : /^mute mic(?:rophone)?\b/i;
    const offPattern = type === 'camera'
      ? /^turn camera on\b/i : /^unmute mic(?:rophone)?\b/i;

    const named = controls.filter(el =>
      onPattern.test(label(el)) || offPattern.test(label(el))
    );
    if (named.length) {
      const el = unique(named);
      return el ? { el, on: onPattern.test(label(el)) } : null;
    }

    const selector = type === 'camera'
      ? '[data-tid="toggle-video"]' : '[data-tid="toggle-mute"]';
    const candidates = [...document.querySelectorAll(selector)]
      .filter(visible).map(el => {
        if (checked(el) !== null) return el;
        return el.querySelector(
          '[role="switch"], input[type="checkbox"], ' +
          '[aria-checked], [aria-pressed]'
        );
      }).filter(el => visible(el) && checked(el) !== null);
    const el = unique([...new Set(candidates)]);
    return el ? { el, on: checked(el) } : null;
  }

  function noAudioSelected() {
    return [...document.querySelectorAll(
      'input[type="radio"], [role="radio"]'
    )].some(el => visible(el) && checked(el) === true &&
      /\b(?:don't|do not) use audio\b/i.test(
        label(el) || normalize(
          [...(el.labels || [])].map(item => item.textContent).join(' ')
        ) || normalize(el.parentElement?.textContent)
      ));
  }

  let active = true;
  let queued = false;
  let running = false;
  let seenRemoval = false;
  let pendingUntil = 0;
  let lastRejoinAt = -Infinity;
  let lastJoinAt = -Infinity;
  let lastReport = '';
  const mediaClicks = new Map();
  const state = {
    version: '3.1.0', active: true, phase: 'starting',
    pending: false, rejoinClicks: 0, joinClicks: 0,
    removalCount: 0, reconnectCount: 0
  };

  const report = phase => {
    state.phase = phase;
    if (phase === lastReport) return;
    lastReport = phase;
    console.info('[Teams Rejoin]', phase);
  };

  const beginAttempt = now => {
    state.pending = true;
    pendingUntil = now + CONFIG.pendingTimeout;
    lastJoinAt = -Infinity;
    mediaClicks.clear();
  };

  function check() {
    if (!active || running) return;
    running = true;

    try {
      const now = Date.now();
      const buttons = [...document.querySelectorAll(
        'button, [role="button"]'
      )].filter(visible);
      const button = name => unique(buttons.filter(el =>
        label(el).toLowerCase() === name.toLowerCase()
      ));

      // Both personal and work/school Teams expose this button.
      // Do not depend on the message being an h1 or having one wording.
      const rejoin = button('Rejoin');
      if (rejoin) {
        if (!seenRemoval) {
          seenRemoval = true;
          state.removalCount++;
          lastRejoinAt = -Infinity;
          beginAttempt(now);
        }
        if (!enabled(rejoin)) {
          report('waiting-for-rejoin-button');
          return;
        }
        if (now - lastRejoinAt < CONFIG.rejoinRetryDelay) return;
        lastRejoinAt = now;
        state.rejoinClicks++;
        report('rejoin-clicked');
        rejoin.click();
        return;
      }
      seenRemoval = false;

      if (buttons.some(el => /^leave(?: meeting)?(?:\s*\(|$)/i
        .test(label(el)))) {
        if (state.pending) state.reconnectCount++;
        state.pending = false;
        mediaClicks.clear();
        report('in-meeting');
        return;
      }

      const join = button('Join now');
      if (!state.pending) {
        if (!CONFIG.joinOnOpen || !join) {
          report('watching');
          return;
        }
        // A one-time prejoin attempt, including after a page reload.
        beginAttempt(now);
      }

      const text = normalize(document.body?.innerText);
      if (/someone will let you in shortly|waiting (?:for .*|in .*lobby)|asked to join/i
        .test(text)) {
        report('waiting-in-lobby');
        return;
      }

      if (now > pendingUntil) {
        // Keep the attempt flagged so joinOnOpen does not restart it forever.
        report('rejoin-timed-out');
        return;
      }

      if (!join) {
        report('waiting-for-teams');
        return;
      }

      const name = document.querySelector(
        '[data-tid="prejoin-display-name-input"], ' +
        'input[placeholder="Type your name"]'
      );
      if (name && !String(name.value || '').trim()) {
        report('enter-your-name-manually');
        return;
      }

      const camera = mediaState('camera');
      const microphone = noAudioSelected()
        ? { el: null, on: false } : mediaState('microphone');

      if (!camera || !microphone) {
        report('cannot-confirm-media-settings');
        return;
      }

      // Click live devices once; observe their off state before joining.
      let waitingForMedia = false;
      for (const [type, device] of [
        ['camera', camera], ['microphone', microphone]
      ]) {
        if (!device.on) {
          mediaClicks.delete(type);
          continue;
        }
        waitingForMedia = true;
        if (!enabled(device.el)) continue;
        if (now - (mediaClicks.get(type) ?? -Infinity) <
          CONFIG.mediaRetryDelay) continue;
        mediaClicks.set(type, now);
        device.el.click();
      }
      if (waitingForMedia) {
        report('waiting-for-media-off');
        return;
      }

      if (!enabled(join)) {
        report('waiting-for-join-button');
        return;
      }
      // Prejoin itself may be inside a dialog in work/school Teams.
      // Only an additional dialog covering prejoin should block the click.
      if ([...document.querySelectorAll(
        '[role="dialog"], [role="alertdialog"]'
      )].some(el => visible(el) && !el.contains(join))) {
        report('waiting-for-dialog');
        return;
      }
      if (now - lastJoinAt < CONFIG.joinRetryDelay) return;

      lastJoinAt = now;
      state.joinClicks++;
      report('join-request-submitted');
      join.click();
    } catch (error) {
      report('error-check-console');
      console.warn('[Teams Rejoin]', error);
    } finally {
      running = false;
    }
  }

  function enqueue() {
    if (!active || queued) return;
    queued = true;
    queueMicrotask(() => {
      queued = false;
      check();
    });
  }

  function onUserClick(event) {
    if (!event.isTrusted) return;
    const target = event.target instanceof Element
      ? event.target.closest('button, [role="button"]') : null;
    if (!target) return;
    const text = label(target);
    if (/^(?:leave|leave meeting|end meeting|end for everyone)(?:\s*\(|$)/i
      .test(text) ||
      (state.pending && /^(?:cancel|dismiss)$/i.test(text))) {
      controller.stop('stopped-after-manual-exit');
    }
  }

  const observer = new MutationObserver(enqueue);
  const timer = setInterval(enqueue, CONFIG.pollInterval);
  const controller = {
    get active() { return active; },
    status: () => ({ ...state }),
    checkNow: check,
    stop(reason = 'stopped') {
      if (!active) return;
      active = false;
      state.active = false;
      state.pending = false;
      observer.disconnect();
      clearInterval(timer);
      document.removeEventListener('click', onUserClick, true);
      document.removeEventListener('visibilitychange', enqueue);
      report(reason);
    }
  };
  window.__teamsAutoRejoin = controller;

  observer.observe(document.documentElement, {
    subtree: true, childList: true, characterData: true,
    attributes: true,
    attributeFilter: [
      'disabled', 'aria-disabled', 'aria-checked', 'aria-pressed',
      'aria-label', 'hidden', 'style', 'class', 'checked'
    ]
  });
  document.addEventListener('click', onUserClick, true);
  document.addEventListener('visibilitychange', enqueue);
  console.info('[Teams Rejoin] Version 3.1.0 active');
  check();
})();