(function () {
    // Nur im Top-Frame aktiv. Der Helper wird ohne Frame-Filter per
    // AddScriptToExecuteOnDocumentCreatedAsync in JEDES Dokument injiziert
    // (siehe WebViewManager.OnCoreWebView2InitializationCompleted) - ohne diesen
    // Guard wuerden F-Tasten-Override und Bild-/Video-Klick-Erkennung auch in
    // eingebetteten iframes greifen (z. B. ein kuenftiges Zahlungs- oder
    // OAuth-iframe von grok.com). Der Host (WebViewManager.ExecuteScriptAsync,
    // CoreWebView2.ExecuteScriptAsync) spricht ohnehin immer den Top-Frame an,
    // daher verlieren __wrokSend & Co. hier nichts.
    if (window.top !== window) return;

    var REQUIRED = __WROK_VERSION__;
    if (window.__wrok && window.__wrok.v >= REQUIRED) return;
    window.__wrok = { v: REQUIRED };

    // Aktivität im WebView an die Host-Anwendung melden.
    const resetActivity = function () { window.chrome.webview.postMessage('resetActivity'); };
    ['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart'].forEach(function (ev) {
        window.addEventListener(ev, resetActivity, { passive: true });
    });

    // -----------------------------------------------------------------------
    // F-Tasten-Override (Makros)
    //
    // Greift nur, solange die Host-Checkbox aktiv ist - dann hat der Host
    // (siehe WebViewManager.ConfigureCore) CoreWebView2.Settings.
    // AreBrowserAcceleratorKeysEnabled=false gesetzt, wodurch Chromiums eigene
    // Behandlung von F1-F12 als "Browser-Accelerator-Keys" wegfaellt und diese
    // Tasten stattdessen hier als normale keydown-Events ankommen. Ist die
    // Checkbox aus, bleibt AreBrowserAcceleratorKeysEnabled=true (Standard) -
    // Chromium faengt F1/F3/F5/F6/F7/F10/F11/F12 dann weiterhin selbst ab,
    // bevor sie ueberhaupt hier ankommen; F2/F4/F8/F9 kommen zwar immer an,
    // werden unten aber durch __wrokFKeyEnabled=false ignoriert.
    // -----------------------------------------------------------------------

    window.__wrokFKeyEnabled = __WROK_FKEY_ENABLED__;
    window.__wrokSetFKeyEnabled = function (v) { window.__wrokFKeyEnabled = !!v; };

    // keyCode 112-123 = F1-F12 (deprecated, aber fuer WebView2/Chromium stabil
    // und einfacher als e.key-String-Vergleiche mit Layout-Sonderfaellen).
    var F_KEYCODES = { 112: 1, 113: 2, 114: 3, 115: 4, 116: 5, 117: 6, 118: 7, 119: 8, 120: 9, 121: 10, 122: 11, 123: 12 };

    window.addEventListener('keydown', function (e) {
        if (!window.__wrokFKeyEnabled) return;
        var n = F_KEYCODES[e.keyCode];
        if (!n) return;
        if (e.altKey) return; // Alt+F4 & Co. nie anfassen.

        if (n === 11) {
            // Fullscreen - manuell, da Chromiums native Behandlung bei
            // deaktivierten Accelerator-Keys wegfaellt. Bleibt unabhaengig vom
            // Makro-Override immer die native F11-Funktion.
            e.preventDefault();
            e.stopPropagation();
            try {
                if (document.fullscreenElement || document.webkitFullscreenElement) {
                    (document.exitFullscreen || document.webkitExitFullscreen).call(document);
                } else {
                    var el = document.documentElement;
                    (el.requestFullscreen || el.webkitRequestFullscreen).call(el);
                }
            } catch (ex) { }
            return;
        }
        if (n === 12) {
            // DevTools - JS kann das Fenster nicht selbst oeffnen, daher an den
            // Host melden (siehe WebViewManager.WebMessageReceived).
            e.preventDefault();
            e.stopPropagation();
            try { window.chrome.webview.postMessage('openDevTools'); } catch (ex) { }
            return;
        }

        // F1-F10 (+ Strg/Umschalt-Kombinationen): Entscheidungslogik (Makro
        // ausloesen, Profil wechseln, Reload) bleibt zentral in C#
        // (MainForm.HandleFKeyOverride, siehe WebViewManager.HandleFKeyMessage).
        e.preventDefault();
        e.stopPropagation();
        try {
            window.chrome.webview.postMessage(
                'fkey:' + n + ':' + (e.ctrlKey ? 1 : 0) + ':' + (e.shiftKey ? 1 : 0) + ':' + (e.altKey ? 1 : 0));
        } catch (ex) { }
    }, true);

    // -----------------------------------------------------------------------
    // Hilfsfunktionen
    // -----------------------------------------------------------------------

    function isProseMirror(el) {
        try {
            if (!el || !el.className) return false;
            var cn = (el.className + '').toString().toLowerCase();
            return cn.indexOf('prosemirror') !== -1 || cn.indexOf('tiptap') !== -1;
        } catch (e) { return false; }
    }

    function tryFocusInput(el) {
        try {
            el.focus();
            if ('setSelectionRange' in el && typeof el.setSelectionRange === 'function') {
                var len = (el.value || '').length;
                try { el.setSelectionRange(len, len); } catch (e) { }
            }
            try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) { }
            try { el.dispatchEvent(new Event('focus', { bubbles: true })); } catch (e) { }
            try { el.scrollIntoView({ block: 'nearest', inline: 'nearest' }); } catch (e) { }
            return true;
        } catch (e) { return false; }
    }

    function placeCaretAtEndContentEditable(el) {
        try {
            el.focus();
            var sel = window.getSelection();
            var range = document.createRange();
            range.selectNodeContents(el);
            range.collapse(false);
            sel.removeAllRanges();
            sel.addRange(range);
            try { el.dispatchEvent(new InputEvent('input', { bubbles: true })); } catch (e) { }
            try { el.dispatchEvent(new Event('focus', { bubbles: true })); } catch (e) { }
            try { el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true })); } catch (e) { }
            try { el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true })); } catch (e) { }
            try { el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true })); } catch (e) { }
            try { el.scrollIntoView({ block: 'nearest', inline: 'nearest' }); } catch (e) { }
            return true;
        } catch (e) { return false; }
    }

    // -----------------------------------------------------------------------
    // __wrokEnsureFocus
    // -----------------------------------------------------------------------

    window.__wrokEnsureFocus = function () {
        try {
            var el = document.activeElement;
            if (!el || el === document.body || !(el.isContentEditable || 'value' in el)) {
                el = document.querySelector('[contenteditable], textarea, input[type=text], input[type=search], [role=textbox]');
            }
            if (!el) {
                el = document.querySelector('textarea, input[type=text], [contenteditable]');
                if (!el) return false;
            }
            var tag = (el.tagName || '').toUpperCase();
            if ((tag === 'INPUT' || tag === 'TEXTAREA' || 'value' in el) && tryFocusInput(el)) return true;
            if (el.isContentEditable && placeCaretAtEndContentEditable(el)) return true;
            var child = el.querySelector('textarea, input[type=text], [contenteditable]');
            if (child) {
                if (child.isContentEditable) return placeCaretAtEndContentEditable(child);
                return tryFocusInput(child);
            }
            try { el.click(); } catch (e) { }
            if (el.isContentEditable) return placeCaretAtEndContentEditable(el);
            return false;
        } catch (e) {
            return false;
        }
    };

    // -----------------------------------------------------------------------
    // __wrokSend
    // -----------------------------------------------------------------------

    window.__wrokSend = function (text, pressEnter) {
        try {
            if (typeof text !== 'string') text = String(text || '');
            var target = document.activeElement;
            if (!target || target === document.body || !(target.isContentEditable || 'value' in target)) {
                target = document.querySelector('[contenteditable], textarea, input[type=text], input[type=search], [role=textbox]');
            }
            if (!target) return false;
            try { if (window.__wrokEnsureFocus) window.__wrokEnsureFocus(); } catch (e) { }
            try { target.focus(); } catch (e) { }
            var tag = (target.tagName || '').toUpperCase();

            // --- Standard-Inputs und Textareas ---
            if (tag === 'INPUT' || tag === 'TEXTAREA' || 'value' in target) {
                var start = typeof target.selectionStart === 'number' ? target.selectionStart : (target.value || '').length;
                var end = typeof target.selectionEnd === 'number' ? target.selectionEnd : start;
                var val = target.value || '';
                var prefix = (start > 0 && val.charAt(start - 1) !== ' ') ? ' ' : '';
                var newVal = val.slice(0, start) + prefix + text + val.slice(end);
                target.value = newVal;
                var newPos = start + prefix.length + text.length;
                try { target.setSelectionRange(newPos, newPos); } catch (e) { }
                try { target.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) { }
                try { target.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) { }
                if (pressEnter) {
                    try {
                        if (target.form) {
                            if (typeof target.form.requestSubmit === 'function') target.form.requestSubmit();
                            else target.form.submit();
                        } else {
                            target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                            target.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                        }
                    } catch (e) { }
                }
                return true;
            }

            // --- ContentEditable ---
            var sel = window.getSelection();
            var range = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
            if (!range) {
                var pfx = (target.innerText && target.innerText.slice(-1) !== ' ') ? ' ' : '';
                target.innerText = (target.innerText || '') + pfx + text;
                var r2 = document.createRange();
                r2.selectNodeContents(target);
                r2.collapse(false);
                sel.removeAllRanges();
                sel.addRange(r2);
                try { target.dispatchEvent(new InputEvent('input', { bubbles: true })); } catch (e) { }
                if (pressEnter) {
                    var btn = document.querySelector('button[type=submit], button[aria-label*="send" i], button[class*="send" i], [role=button][aria-label*="send" i]');
                    if (btn) { try { btn.click(); } catch (e) { } }
                    else {
                        try { target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) { }
                        try { target.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) { }
                    }
                }
                return true;
            }

            // Cursor-Position ermitteln und Leerzeichen-Prefix einfügen falls nötig.
            var insertPrefix = '';
            var sc = range.startContainer;
            var off = range.startOffset;
            var prevChar = '';
            if (sc.nodeType === Node.TEXT_NODE) {
                if (off > 0) prevChar = sc.textContent.charAt(off - 1) || '';
                else {
                    var prev = sc.previousSibling;
                    if (prev && prev.nodeType === Node.TEXT_NODE)
                        prevChar = prev.textContent.charAt(prev.textContent.length - 1) || '';
                }
            } else {
                var prevNode = range.startContainer.childNodes[off - 1];
                if (prevNode && prevNode.nodeType === Node.TEXT_NODE)
                    prevChar = prevNode.textContent.charAt(prevNode.textContent.length - 1) || '';
            }
            if (prevChar && prevChar !== ' ') insertPrefix = ' ';

            var node = document.createTextNode(insertPrefix + text);
            range.insertNode(node);
            range.setStartAfter(node);
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);
            try { target.dispatchEvent(new InputEvent('input', { bubbles: true })); } catch (e) { }

            if (pressEnter) {
                var btn2 = document.querySelector('button[type=submit], button[aria-label*="send" i], button[class*="send" i], [role=button][aria-label*="send" i]');
                if (btn2) { try { btn2.click(); } catch (e) { } }
                else {
                    try { range.insertNode(document.createElement('br')); } catch (e) { }
                    try { range.setStartAfter(node.nextSibling || node); } catch (e) { }
                    try { range.collapse(true); } catch (e) { }
                    try { sel.removeAllRanges(); sel.addRange(range); } catch (e) { }
                    try { target.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) { }
                    try { target.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true })); } catch (e) { }
                }
            }
            return true;
        } catch (e) {
            return false;
        }
    };
    // -----------------------------------------------------------------------
    // Bild-Klick → openImage-Nachricht an Host
    // -----------------------------------------------------------------------

    (function () {
        // Icons und Avatare ignorieren – nur echte Inhaltsbilder öffnen.
        var MIN_SIZE = 64;

        // Kein Inhaltsbild: Cookie-Banner-Logos, Profilbilder, Avatare, Emoji-Sprites.
        var IGNORE_URL = /cookielaw\.org|onetrust|profile-picture|\/avatars?\/|\/emoji\//i;

        function isContentUrl(url) {
            return !!url && url.indexOf('data:') !== 0 && !IGNORE_URL.test(url);
        }

        // Interaktive Elemente von der (unspezifischeren) CSS-Hintergrundbild-Erkennung
        // ausschliessen - ein Button/Link/Formularelement, das nur dekorativ per
        // background-image gestaltet ist (statt eines <img>), soll beim Klick seine
        // eigentliche Funktion ausloesen und nicht faelschlich den Bildbetrachter
        // oeffnen. Gilt bewusst NICHT fuer <img>/<video> selbst - ein Bild innerhalb
        // eines Links/Buttons soll weiterhin normal oeffnen.
        var INTERACTIVE_SEL = 'button, [role="button"], a, input, select, textarea, ' +
            '[contenteditable="true"], [role="link"], [role="menuitem"], [role="tab"], ' +
            '[role="checkbox"], [role="radio"], [role="switch"]';

        function isInteractive(node) {
            try { return typeof node.matches === 'function' && node.matches(INTERACTIVE_SEL); }
            catch (e) { return false; }
        }

        // Liefert { url, kind } eines Elements – kind ist 'image' oder 'video'.
        // Kein URL-Pattern-Matching auf Dateiendungen, da Grok CDN-Pfade ohne
        // Endung verwendet (z. B. .../projects/<id>/<id>/asset).
        function mediaOf(node) {
            if (!node || node.nodeType !== 1) return null;
            try {
                if (node.tagName === 'VIDEO') {
                    var vurl = node.currentSrc || node.src || node.getAttribute('src') || '';
                    if (!vurl) {
                        var src = node.querySelector('source');
                        if (src) vurl = src.src || src.getAttribute('src') || '';
                    }
                    return isContentUrl(vurl) ? { url: vurl, kind: 'video' } : null;
                }
                if (node.tagName === 'IMG') {
                    var w = node.naturalWidth || node.width || 0;
                    var h = node.naturalHeight || node.height || 0;
                    if (w && h && (w < MIN_SIZE || h < MIN_SIZE)) return null;
                    var url = node.currentSrc || node.src || node.getAttribute('src') || '';
                    return isContentUrl(url) ? { url: url, kind: 'image' } : null;
                }
                if (isInteractive(node)) return null;

                var r = node.getBoundingClientRect();
                if (r.width < MIN_SIZE || r.height < MIN_SIZE) return null;
                var bg = window.getComputedStyle(node).backgroundImage;
                if (bg && bg !== 'none') {
                    var m = /url\(["']?([^"')]+)["']?\)/.exec(bg);
                    if (m && isContentUrl(m[1])) return { url: m[1], kind: 'image' };
                }
            } catch (e) { }
            return null;
        }

        function findMedia(e) {
            // 1) Vom Klickziel nach oben durch die Elternkette.
            var node = e.target;
            for (var i = 0; i < 6 && node && node !== document.documentElement; i++) {
                var hit = mediaOf(node);
                if (hit) return hit;
                node = node.parentElement;
            }
            // 2) Fallback: alle Elemente unter dem Mauszeiger. Nötig, weil Grok ein
            //    transparentes Overlay über die Medien legt – dann ist e.target das
            //    Overlay und das <img>/<video> liegt DARUNTER, nicht in der Elternkette.
            try {
                var stack = document.elementsFromPoint(e.clientX, e.clientY) || [];
                for (var j = 0; j < stack.length; j++) {
                    var hit2 = mediaOf(stack[j]);
                    if (hit2) return hit2;
                }
            } catch (ex) { }
            return null;
        }

        document.addEventListener('click', function (e) {
            try {
                var hit = findMedia(e);
                if (!hit) return;
                // Kein preventDefault/stopPropagation — Grok soll sein Overlay
                // weiterhin normal öffnen und schließen können.
                window.chrome.webview.postMessage(
                    (hit.kind === 'video' ? 'openVideo:' : 'openImage:') + hit.url);
            } catch (ex) {}
        }, true);
    })();

    // -----------------------------------------------------------------------
    // Async-Brücke zum Host
    //
    // WICHTIG: CoreWebView2.ExecuteScriptAsync löst KEINE Promises auf – ein
    // zurückgegebenes Promise wird als "{}" serialisiert. Asynchrone Ergebnisse
    // müssen deshalb per postMessage zurückgemeldet werden.
    // Format: wrokResult:<token>:<payload>   (leerer payload = Fehler/kein Ergebnis)
    // -----------------------------------------------------------------------

    function postResult(token, payload) {
        try {
            window.chrome.webview.postMessage(
                'wrokResult:' + token + ':' + (payload == null ? '' : payload));
        } catch (e) { }
    }

    // Bild über die eingeloggte Session laden und als Base64 zurückgeben.
    window.__wrokFetchImage = function (url, token) {
        try {
            fetch(url, { credentials: 'include' })
                .then(function (r) { return r.ok ? r.arrayBuffer() : null; })
                .then(function (buf) {
                    if (!buf) { postResult(token, ''); return; }
                    // In Blöcken konvertieren – String.fromCharCode.apply hat ein
                    // Argumentlimit, und Zeichen-für-Zeichen ist bei großen Bildern zu langsam.
                    var bytes = new Uint8Array(buf), bin = '', CHUNK = 0x8000;
                    for (var i = 0; i < bytes.length; i += CHUNK)
                        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
                    postResult(token, btoa(bin));
                })
                .catch(function () { postResult(token, ''); });
        } catch (e) { postResult(token, ''); }
    };

    // Content-Type einer URL ermitteln (für den Zwischenablage-Weg, da Grok-URLs
    // keine Dateiendung haben). HEAD zuerst, sonst 1-Byte-Range-GET als Rückfall.
    window.__wrokProbeType = function (url, token) {
        function done(r) {
            postResult(token, (r && r.ok && r.headers.get('content-type')) || '');
        }
        try {
            fetch(url, { method: 'HEAD', credentials: 'include' })
                .then(function (r) {
                    if (r.ok && r.headers.get('content-type')) { done(r); return; }
                    return fetch(url, {
                        method: 'GET', credentials: 'include',
                        headers: { 'Range': 'bytes=0-0' }
                    }).then(done);
                })
                .catch(function () {
                    fetch(url, {
                        method: 'GET', credentials: 'include',
                        headers: { 'Range': 'bytes=0-0' }
                    }).then(done).catch(function () { postResult(token, ''); });
                });
        } catch (e) { postResult(token, ''); }
    };

    // Rate-Limits der angegebenen Modelle abfragen.
    window.__wrokRateLimits = function (token, models) {
        function fetchLimit(modelName) {
            return fetch('https://grok.com/rest/rate-limits', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ requestKind: 'DEFAULT', modelName: modelName })
            }).then(function (r) {
                if (r.status === 401 || r.status === 403) return { error: 'UNAUTHORIZED' };
                if (!r.ok) return { error: 'HTTP ' + r.status };
                return r.json();
            }).catch(function (e) {
                return { error: String((e && e.message) || e) };
            });
        }
        try {
            var list = models && models.length ? models : ['grok-3', 'grok-4-heavy'];
            Promise.all(list.map(fetchLimit)).then(function (res) {
                var out = {};
                for (var i = 0; i < list.length; i++) out[list[i]] = res[i];
                postResult(token, JSON.stringify(out));
            }).catch(function () { postResult(token, ''); });
        } catch (e) { postResult(token, ''); }
    };

})();
