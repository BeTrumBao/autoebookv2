// ==UserScript==
// @name         AutoEbook
// @namespace    http://tampermonkey.net/
// @version      2026.1.6
// @description  AutoEbook Made By Em Wizi With Love <3
// @author       Em Wizi
// @match        *://thcscaukieu.thuvienhcm.edu.vn/*
// @match        https://*.officeapps.live.com/*
// @updateURL    https://raw.githubusercontent.com/BeTrumBao/autoebook/refs/heads/main/script_v2.js
// @downloadURL  https://raw.githubusercontent.com/BeTrumBao/autoebook/refs/heads/main/script_v2.js
// @grant        none
// ==/UserScript==

(function () {
    'use strict';
    const APP_TAG = '[AutoEbooks]';
    const _d = (s) => {
        try {
            return decodeURIComponent(escape(atob(s)));
        } catch {
            return s;
        }
    };

    const isOfficeAppsHost = /(^|\.)officeapps\.live\.com$/i.test(location.hostname);
    const isOfficeWordHost = /(?:^|-)word-view\.officeapps\.live\.com$/i.test(location.hostname);
    if (isOfficeAppsHost) {
        if (isOfficeWordHost && window.top !== window) {
            let announced = false;
            let markerSeenAt = 0;
            const hasVisibleDocumentSurface = () => {
                const selectors = '[role="document"], canvas, [id*="WACViewPanel"], [class*="WACViewPanel"], [class*="DocumentPage"], [class*="document-page"]';
                return Array.from(document.querySelectorAll(selectors)).some((el) => {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 80 && rect.height > 80;
                });
            };
            const announceWhenReady = () => {
                if (announced || document.readyState !== 'complete' || !document.getElementById('SharedAuthFrame')) return;
                if (!markerSeenAt) markerSeenAt = Date.now();
                if (Date.now() - markerSeenAt < 1800) return;
                if (!hasVisibleDocumentSurface() && (document.body?.innerText || '').trim().length < 120) return;
                announced = true;
                window.top.postMessage({ type: 'AUTOEBOOKS_WORD_DOCUMENT_READY', marker: 'SharedAuthFrame' }, '*');
            };
            const poll = setInterval(() => {
                announceWhenReady();
                if (announced) clearInterval(poll);
            }, 500);
            window.addEventListener('load', announceWhenReady);
            window.addEventListener('message', (event) => {
                if (event.source === window.top && event.data?.type === 'AUTOEBOOKS_WORD_DOCUMENT_READY_ACK') {
                    clearInterval(poll);
                }
            });
            announceWhenReady();
        }
        return;
    }

    const isOfficeWordFrame = (frame) => {
        try { return /(?:^|-)word-view\.officeapps\.live\.com$/i.test(new URL(frame.src, location.href).hostname); }
        catch (e) { return false; }
    };
    let wordReaderReadyFrame = null;
    let wordReaderReadyAt = 0;
    window.addEventListener('message', (event) => {
        if (event.data?.type !== 'AUTOEBOOKS_WORD_DOCUMENT_READY') return;
        const frame = Array.from(document.querySelectorAll('iframe')).find((item) => item.contentWindow === event.source && isOfficeWordFrame(item));
        const trustedWordOrigin = /^https:\/\/[a-z0-9-]*word-view\.officeapps\.live\.com$/i.test(event.origin);
        if (!frame && !trustedWordOrigin) return;
        if (frame && wordReaderReadyFrame !== frame) {
            wordReaderReadyFrame = frame;
        }
        wordReaderReadyAt = Date.now();
        event.source.postMessage({ type: 'AUTOEBOOKS_WORD_DOCUMENT_READY_ACK' }, '*');
    });

    if (window.__THUVIEN_TOOL_LOADED__) return;
    window.__THUVIEN_TOOL_LOADED__ = true;

    try {
        const WORKER_CACHE_KEY = 'tr_pdf_worker_cache_v1';
        let cachedWorkerBlobUrl = null;

        // Khởi tạo Blob URL từ cache đồng bộ (localStorage) nếu có sẵn
        try {
            const cachedCode = localStorage.getItem(WORKER_CACHE_KEY);
            if (cachedCode && cachedCode.length > 5000) {
                const blob = new Blob([cachedCode], { type: 'application/javascript' });
                cachedWorkerBlobUrl = URL.createObjectURL(blob);
            }
        } catch (e) { }

        // IndexedDB helper để sao lưu/phục hồi nếu localStorage bị giới hạn dung lượng
        const DB_NAME = 'AutoEbooks_DB';
        const DB_STORE = 'worker_store';
        const openWorkerDB = () => new Promise((resolve, reject) => {
            if (!window.indexedDB) return reject('No indexedDB');
            const req = indexedDB.open(DB_NAME, 1);
            req.onupgradeneeded = () => {
                const db = req.result;
                if (!db.objectStoreNames.contains(DB_STORE)) {
                    db.createObjectStore(DB_STORE);
                }
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });

        const saveWorkerToDB = async (code) => {
            try {
                const db = await openWorkerDB();
                const tx = db.transaction(DB_STORE, 'readwrite');
                tx.objectStore(DB_STORE).put(code, 'pdf_worker_code');
            } catch (e) { }
        };

        const loadWorkerFromDB = async () => {
            try {
                const db = await openWorkerDB();
                const tx = db.transaction(DB_STORE, 'readonly');
                const req = tx.objectStore(DB_STORE).get('pdf_worker_code');
                req.onsuccess = () => {
                    const code = req.result;
                    if (code && code.length > 5000 && !cachedWorkerBlobUrl) {
                        const blob = new Blob([code], { type: 'application/javascript' });
                        cachedWorkerBlobUrl = URL.createObjectURL(blob);
                        try { localStorage.setItem(WORKER_CACHE_KEY, code); } catch (e) { }
                    }
                };
            } catch (e) { }
        };
        if (!cachedWorkerBlobUrl) loadWorkerFromDB();

        const isPdfWorkerUrl = (url) => {
            if (!url || typeof url !== 'string') return false;
            return /pdf\.worker|pdfjs.*worker|\.worker\.js/i.test(url);
        };

        let isCachingWorker = false;
        const cacheWorkerFromUrl = (url) => {
            if (isCachingWorker || !url || url.startsWith('blob:')) return;
            isCachingWorker = true;
            fetch(url)
                .then(res => {
                    if (!res.ok) throw new Error('HTTP ' + res.status);
                    return res.text();
                })
                .then(code => {
                    if (code && code.length > 5000) {
                        try { localStorage.setItem(WORKER_CACHE_KEY, code); } catch (e) { }
                        saveWorkerToDB(code);
                        if (!cachedWorkerBlobUrl) {
                            const blob = new Blob([code], { type: 'application/javascript' });
                            cachedWorkerBlobUrl = URL.createObjectURL(blob);
                        }
                    }
                })
                .catch(e => console.warn(`${APP_TAG} Không thể tải cache PDF Worker:`, e))
                .finally(() => { isCachingWorker = false; });
        };

        // Hook Worker để nạp thẳng từ Blob URL cục bộ
        const OrigWorker = window.Worker;
        window.Worker = function (scriptURL, options) {
            try {
                const urlStr = typeof scriptURL === 'string' ? scriptURL : (scriptURL?.toString ? scriptURL.toString() : '');
                if (isPdfWorkerUrl(urlStr)) {
                    if (cachedWorkerBlobUrl) {
                        return new OrigWorker(cachedWorkerBlobUrl, options);
                    } else {
                        cacheWorkerFromUrl(urlStr);
                    }
                }
            } catch (e) {
                console.warn(`${APP_TAG} Lỗi hook Worker:`, e);
            }
            return new OrigWorker(scriptURL, options);
        };
        window.Worker.prototype = OrigWorker.prototype;

        const applyDFlip2D = (obj) => {
            if (!obj) return;
            try {
                if (obj.defaults) {
                    obj.defaults.webgl = false;
                    obj.defaults.mode = '2d';
                    if (cachedWorkerBlobUrl) {
                        obj.defaults.pdfWorkerSrc = cachedWorkerBlobUrl;
                    }
                }
            } catch (e) { }
        };
        if (window.DFLIP) applyDFlip2D(window.DFLIP);
        let _rawDFLIP = window.DFLIP;
        Object.defineProperty(window, 'DFLIP', {
            configurable: true,
            enumerable: true,
            get() { return _rawDFLIP; },
            set(v) {
                _rawDFLIP = v;
                applyDFlip2D(_rawDFLIP);
            }
        });

        const origGetContext = HTMLCanvasElement.prototype.getContext;
        HTMLCanvasElement.prototype.getContext = function (type, attributes) {
            if (typeof type === 'string' && (type === 'webgl' || type === 'webgl2' || type === 'experimental-webgl')) {
                return null;
            }
            return origGetContext.call(this, type, attributes);
        };

        const FAST_PDF_URL = 'https://raw.githubusercontent.com/BeTrumBao/autoebook/refs/heads/main/data_pdf_ez.pdf';
        const isHeavyPdfUrl = (url) => {
            if (!url || typeof url !== 'string') return false;
            if (url.includes('data_pdf_ez.pdf')) return false;
            return url.toLowerCase().includes('.pdf') || url.includes('digitaloceanspaces.com');
        };

        const origFetch = window.fetch;
        window.fetch = function (input, init) {
            try {
                if (typeof input === 'string' && isHeavyPdfUrl(input)) {
                    return origFetch.call(this, FAST_PDF_URL, init);
                } else if (input instanceof Request && isHeavyPdfUrl(input.url)) {
                    const newReq = new Request(FAST_PDF_URL, input);
                    return origFetch.call(this, newReq, init);
                }
            } catch (e) {
                console.warn('[AutoEbooks] Lỗi hook fetch PDF:', e);
            }
            return origFetch.apply(this, arguments);
        };

        const origXHROpen = XMLHttpRequest.prototype.open;
        XMLHttpRequest.prototype.open = function (method, url) {
            try {
                if (typeof url === 'string' && isHeavyPdfUrl(url)) {
                    arguments[1] = FAST_PDF_URL;
                }
            } catch (e) {
                console.warn('[AutoEbooks] Lỗi hook XHR PDF:', e);
            }
            return origXHROpen.apply(this, arguments);
        };

    } catch (e) {
        console.warn('[AutoEbooks] Không thể inject WebGL PDF Hook:', e);
    }

    const KEY_ACTIVE = 'tr_auto_active';
    const KEY_LIST = 'tr_auto_list';
    const KEY_INDEX = 'tr_auto_index';
    const KEY_SCRAPE_PENDING = 'tr_scrape_pending';
    const KEY_READ_HISTORY = 'tr_read_history_v16';
    const KEY_STOP = 'tr_auto_stop';
    const KEY_SKIP_CURRENT = 'tr_auto_skip_current';
    const KEY_UI_COLLAPSED = 'tr_ui_collapsed';
    const KEY_SCAN_ACTIVE = 'tr_scan_active';
    const KEY_SCAN_LIST = 'tr_scan_list';
    const KEY_SCAN_INDEX = 'tr_scan_index';
    const KEY_SCAN_RESULT = 'tr_scan_result';
    const KEY_TIMING = 'tr_timing_stats_v1';
    const KEY_LIMIT = 'tr_book_limit';
    const KEY_BG_URL = 'tr_bg_url';
    const KEY_PDF_TIMES = 'tr_pdf_load_times_v1';
    const KEY_RETRY = 'tr_retry_count_v1';
    const KEY_AUTO_MODE = 'tr_auto_mode';
    const KEY_VIEW_COUNT_MODE = 'tr_view_count_mode';
    const KEY_VIEW_LOOP_TOTAL = 'tr_view_loop_total';
    const KEY_VIEW_LOOP_CURRENT = 'tr_view_loop_current';
    const KEY_SKIP_COUNT = 'tr_skip_count';
    const KEY_TEACHER_COUNT = 'tr_teacher_count';
    const KEY_TARGET_COUNT = 'tr_target_count';
    const KEY_READ_SUCCESS_COUNT = 'tr_read_success_count';
    const KEY_SKIPPED_DETAILS = 'tr_skipped_details';
    const KEY_PERF_MODE = 'tr_perf_mode';
    const KEY_BOOK_CACHE = 'tr_book_cache_v1';
    const KEY_BOOK_CACHE_TIME = 'tr_book_cache_time_v1';
    const KEY_BOOK_PACK_URL = 'tr_book_pack_url';
    const KEY_RESOURCE_VER = 'tr_resource_ver';
    const DEFAULT_BOOK_PACK_URL = 'https://raw.githubusercontent.com/BeTrumBao/autoebook/refs/heads/main/books_pack.json';
    const KEY_ANIM_OFF = 'tr_anim_off';
    const KEY_FAST_MODE = 'tr_fast_read_mode';
    const KEY_START_TIME = 'tr_session_start_time';
    const CACHE_TTL_MS = 30 * 60 * 1000;

    const DEFAULT_SEC_PER_BOOK = 25;
    const MAX_RETRY_PER_URL = 2;
    const LIST_URL = 'https://thcscaukieu.thuvienhcm.edu.vn/tai-lieu/moi-nhat';
    const DEFAULT_BG = 'https://theselfishmeme.co.uk/wp-content/uploads/2025/11/avatar-anime-nu-sang-cute-ngau-14.webp';

    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const isStopped = () => localStorage.getItem(KEY_STOP) === 'true' || localStorage.getItem(KEY_SKIP_CURRENT) === 'true';
    const isPerfMode = () => localStorage.getItem(KEY_PERF_MODE) === 'true';
    const isAnimOff = () => false;
    const isFastMode = () => true;

    const generateBookUID = (url, title = '') => {
        let h = 2166136261 >>> 0;
        const str = (url || title || '').trim();
        for (let i = 0; i < str.length; i++) {
            h ^= str.charCodeAt(i);
            h = Math.imul(h, 16777619) >>> 0;
        }
        return h;
    };


    const clearBookCache = () => {
        localStorage.removeItem(KEY_BOOK_CACHE);
        localStorage.removeItem(KEY_BOOK_CACHE_TIME);
        localStorage.removeItem('tr_pack_loaded');
    };

    const getFooterNoteText = () => {
        const resVer = localStorage.getItem(KEY_RESOURCE_VER);
        return resVer
            ? `Auto Ebooks - 2026.1.6 | Res: ${resVer} - Made By Em Wizi With Love <3`
            : 'Auto Ebooks - 2026.1.6 - Made By Em Wizi With Love <3';
    };

    const fetchBookPackFromGithub = async (updateStatus) => {
        const configuredUrl = localStorage.getItem(KEY_BOOK_PACK_URL);
        const candidateUrls = configuredUrl
            ? [configuredUrl]
            : [
                'https://raw.githubusercontent.com/BeTrumBao/autoebook/refs/heads/main/book_packs.json',
                'https://raw.githubusercontent.com/BeTrumBao/autoebook/refs/heads/main/books_pack.json'
            ];

        for (const packUrl of candidateUrls) {
            try {
                if (typeof updateStatus === 'function') updateStatus(_d('IMSQYW5nIHThuqNpIGfDs2kgdMOgaSBuZ3V5w6puIHPDoWNoLi4u'), 'loading');
                const urlWithTs = packUrl + (packUrl.includes('?') ? '&' : '?') + '_t=' + Date.now();
                const res = await fetch(urlWithTs, { cache: 'no-cache' });
                if (!res.ok) continue;
                const data = await res.json();
                const list = Array.isArray(data) ? data : (Array.isArray(data?.books) ? data.books : null);
                if (!list || list.length === 0) continue;

                const resVer = data?.metadata?.resource_ver || data?.metadata?.version || '';
                if (resVer) {
                    localStorage.setItem(KEY_RESOURCE_VER, resVer);
                    const fnEl = document.querySelector('.tr-footer-note');
                    if (fnEl) fnEl.textContent = getFooterNoteText();
                }

                const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
                const formatted = list.map(b => ({
                    uid: b.uid || generateBookUID(b.url, b.title),
                    title: b.title || 'Không tên',
                    url: b.url,
                    isRead: Boolean(b.isRead) || history.includes(b.url)
                })).filter(b => b.url);

                if (formatted.length > 0) {

                    if (typeof updateStatus === 'function') updateStatus(_d('IMSQw6MgbuG6oXAg') + formatted.length + _d('IHPDoWNoIHThu6sgR8OzaSB0w6BpIG5ndXnDqm4h'), 'success');
                    return formatted;
                }
            } catch (e) {
                console.warn('[AutoEbooks] Chưa nạp được gói tài nguyên sách:', e);
            }
        }
        return null;
    };

    const showMaterialBlueDialog = (title, message, okText = 'Đã hiểu') => new Promise((resolve) => {
        const old = document.getElementById('tr-blue-alert-overlay');
        if (old) old.remove();

        const overlay = document.createElement('div');
        overlay.id = 'tr-blue-alert-overlay';
        const box = document.createElement('div');
        box.className = 'tr-blue-alert-box';
        box.innerHTML = `
            <div class="tr-blue-alert-icon">
                <svg viewBox="0 0 24 24" width="30" height="30" fill="currentColor">
                    <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
                </svg>
            </div>
            <div class="tr-blue-alert-title">${title}</div>
            <div class="tr-blue-alert-body">${message}</div>
            <div class="tr-blue-alert-footer">
                <button class="tr-btn tr-btn-primary tr-blue-alert-btn" id="tr-blue-alert-ok">${okText}</button>
            </div>
        `;
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        const close = () => { overlay.remove(); resolve(true); };
        box.querySelector('#tr-blue-alert-ok').addEventListener('click', close);
        overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    });

    const hasAccessToken = () => {
        try {
            const match = document.cookie.match(/(?:^|;\s*)access_token=([^;]+)/);
            if (match && match[1] && match[1].trim() !== '') return true;
            if (localStorage.getItem('access_token') || sessionStorage.getItem('access_token')) return true;
        } catch { }
        return false;
    };

    const checkLoginOrAlert = async (updateStatus) => {
        if (!hasAccessToken()) {
            await showMaterialBlueDialog(
                'Chưa đăng nhập tài khoản',
                'Bạn chưa đăng nhập tài khoản thư viện. Vui lòng đăng nhập trước khi sử dụng tính năng!',
                'Đã hiểu'
            );
            if (typeof updateStatus === 'function') {
                updateStatus(_d('IEzhu5dpOiBDaMawYSDEkcSDbmcgbmjhuqFwIHTDoGkga2hv4bqjbiE='), 'error');
            }
            return false;
        }
        return true;
    };

    const getOrFetchBooks = async (updateStatus) => {
        const packBooks = await fetchBookPackFromGithub(updateStatus);
        if (packBooks && packBooks.length > 0) return packBooks;

        await showMaterialBlueDialog(
            'Lỗi tải gói tài nguyên sách',
            'Không thể tải dữ liệu gói sách. Vui lòng kiểm tra lại kết nối mạng!',
            'Đã hiểu'
        );
        if (typeof updateStatus === 'function') updateStatus(_d('IEzhu5dpOiBLaMO0bmcgdGjhu4MgdOG6o2kgZ8OzaSB0w6BpIG5ndXnDqm4gc8OhY2gh'), 'error');
        return null;
    };

    const applyAnimOff = () => {
        const id = 'tr-anim-off-style';
        if (isAnimOff()) {
            if (!document.getElementById(id)) {
                const style = document.createElement('style');
                style.id = id;
                style.textContent = `
                        *, *::before, *::after {
                            animation-duration: 0s !important;
                            animation-delay: 0s !important;
                            transition-duration: 0s !important;
                            transition-delay: 0s !important;
                        }
                        .tr-panel, .tr-panel *, .tr-toast, #tr-scrape-overlay, #tr-scrape-overlay * {
                            animation-duration: 0s !important;
                            transition-duration: 0s !important;
                        }
                    `;
                document.head.appendChild(style);
            }
        } else {
            const el = document.getElementById(id);
            if (el) el.remove();
        }
    };
    const isListPage = (url = window.location.href) =>
        url.includes('/tai-lieu/moi-nhat') ||
        url.endsWith('/tai-lieu') ||
        url.endsWith('/tai-lieu/');

    const fmtTime = (sec) => {
        if (sec < 0 || !isFinite(sec)) return '--';
        sec = Math.round(sec);
        if (sec < 60) return sec + ' giây';
        const m = Math.floor(sec / 60);
        const s = sec % 60;
        if (m < 60) return m + ' phút ' + (s > 0 ? s + ' giây' : '');
        const h = Math.floor(m / 60);
        const mm = m % 60;
        return h + ' giờ ' + (mm > 0 ? mm + ' phút' : '');
    };
    const fmtClock = (d = new Date()) => {
        const p = (n) => String(n).padStart(2, '0');
        return p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
    };

    const getBgUrl = () => localStorage.getItem(KEY_BG_URL) || DEFAULT_BG;
    const testImage = (url) => new Promise((resolve) => {
        const img = new Image(); let done = false;
        img.onload = () => { if (!done) { done = true; resolve(true); } };
        img.onerror = () => { if (!done) { done = true; resolve(false); } };
        img.src = url;
        setTimeout(() => { if (!done) { done = true; resolve(false); } }, 5000);
    });

    const showErrorAlert = (msg) => {
        try {
            if (localStorage.getItem(KEY_ACTIVE) === 'true' || localStorage.getItem(KEY_SCAN_ACTIVE) === 'true') {
                console.warn('[AutoEbooks] Error:', msg);
                showToast(msg, '');
                return;
            }
            window.alert(' [Lỗi]\n' + msg);
        } catch (e) {
            console.error('[AutoEbooks] Alert error:', e);
        }
    };

    const showToast = (message) => {
        let old = document.getElementById('tr-toast');
        if (old) old.remove();
        const toast = document.createElement('div');
        toast.id = 'tr-toast';
        toast.className = 'tr-toast';
        toast.textContent = message;
        document.body.appendChild(toast);
        setTimeout(() => toast.classList.add('tr-toast-show'), 15);
        setTimeout(() => {
            toast.classList.remove('tr-toast-show');
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    };

    const getRetryState = () => {
        try { return JSON.parse(localStorage.getItem(KEY_RETRY) || '{}'); }
        catch (e) { return {}; }
    };
    const getRetryCount = (url) => {
        const s = getRetryState();
        return s[url] || 0;
    };
    const bumpRetry = (url) => {
        const s = getRetryState();
        s[url] = (s[url] || 0) + 1;
        const keys = Object.keys(s);
        if (keys.length > 30) delete s[keys[0]];
        try { localStorage.setItem(KEY_RETRY, JSON.stringify(s)); } catch (e) { }
        return s[url];
    };
    const clearRetry = (url) => {
        const s = getRetryState();
        if (s[url]) { delete s[url]; try { localStorage.setItem(KEY_RETRY, JSON.stringify(s)); } catch (e) { } }
    };

    const hasFileAccessError = () => false;

    const getPdfTimes = () => {
        try {
            const raw = JSON.parse(localStorage.getItem(KEY_PDF_TIMES) || '[]');
            // Lọc bỏ các mốc thời gian của sách thật cũ (> 8000ms) để không kéo dài timeout của sách ảo
            return raw.filter(t => typeof t === 'number' && t > 0 && t <= 8000);
        }
        catch (e) { return []; }
    };
    const recordPdfTime = (ms) => {
        if (ms < 50 || ms > 60000) return;
        const arr = getPdfTimes();
        arr.push(ms);
        while (arr.length > 20) arr.shift();
        try { localStorage.setItem(KEY_PDF_TIMES, JSON.stringify(arr)); } catch (e) { }
    };
    const getPdfTimeout = () => {
        const arr = getPdfTimes();
        if (arr.length < 3) return 4000;
        const sorted = [...arr].sort((a, b) => a - b);
        const p90 = sorted[Math.floor(sorted.length * 0.9)] || sorted[sorted.length - 1];
        return Math.max(3000, Math.min(8000, Math.round(p90 * 1.5)));
    };

    const getTimingStats = () => {
        try {
            const s = JSON.parse(localStorage.getItem(KEY_TIMING) || '{}');
            return { samples: Array.isArray(s.samples) ? s.samples : [], avgMs: typeof s.avgMs === 'number' ? s.avgMs : 0 };
        } catch (e) { return { samples: [], avgMs: 0 }; }
    };
    const recordBookTime = (durationMs) => {
        if (durationMs < 3000 || durationMs > 300000) return;
        const stats = getTimingStats();
        stats.samples.push(durationMs);
        while (stats.samples.length > 30) stats.samples.shift();
        const sum = stats.samples.reduce((a, b) => a + b, 0);
        stats.avgMs = Math.round(sum / stats.samples.length);
        try { localStorage.setItem(KEY_TIMING, JSON.stringify(stats)); } catch (e) { }
    };
    const getAvgSecPerBook = () => {
        if (isFastMode()) return 3.5;
        const stats = getTimingStats();
        return stats.avgMs > 0 ? stats.avgMs / 1000 : DEFAULT_SEC_PER_BOOK;
    };

    const getLimit = () => {
        const n = parseInt(localStorage.getItem(KEY_LIMIT) || '0', 10);
        return isNaN(n) || n < 0 ? 0 : n;
    };
    const setLimit = (n) => {
        const v = Math.max(0, parseInt(n, 10) || 0);
        localStorage.setItem(KEY_LIMIT, v.toString());
        return v;
    };

    const mergeScanResults = () => {
        try {
            const scanResults = JSON.parse(localStorage.getItem(KEY_SCAN_RESULT) || '[]');
            if (scanResults.length === 0) return 0;
            const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
            let added = 0;
            scanResults.forEach(r => {
                if (r.status === 'done' && !history.includes(r.url)) { history.push(r.url); added++; }
            });
            if (added > 0) localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
            return added;
        } catch (e) { return 0; }
    };

    const isBookSlider100 = () => {
        try {
            const tracks = document.querySelectorAll('.MuiSlider-track');
            for (const tr of tracks) {
                const w = tr.style.width || '';
                if (w === '100%' || w.startsWith('100')) return true;
                const cssText = tr.style.cssText || '';
                if (cssText.includes('width: 100%') || cssText.includes('width:100%')) return true;
            }
            const prog = getBookProgress();
            if (prog !== null && prog >= 100) return true;
            return false;
        } catch (e) {
            return false;
        }
    };

    const getBookProgress = () => {
        try {
            const labelBox = document.getElementById('label-box');
            if (labelBox) {
                const root = labelBox.closest('.MuiSlider-root');
                if (root) {
                    const inp = root.querySelector('input[type="range"]');
                    if (inp) {
                        const v = parseInt(inp.getAttribute('aria-valuenow') || inp.value || '0', 10);
                        if (!isNaN(v)) return v;
                    }
                }
            }
            const inp = document.querySelector('.MuiSlider-root input[type="range"]');
            if (inp) {
                const v = parseInt(inp.getAttribute('aria-valuenow') || inp.value || '0', 10);
                if (!isNaN(v)) return v;
            }
            const tr = document.querySelector('.MuiSlider-track');
            if (tr && tr.style && tr.style.width) {
                const match = tr.style.width.match(/(\d+)%/);
                if (match) return parseInt(match[1], 10);
            }
            return null;
        } catch (e) { return null; }
    };

    const findReadButton = () => {
        const btns = Array.from(document.querySelectorAll('button'));
        for (const b of btns) {
            if (b.closest('#tr-panel') || b.closest('#tr-view-scrape-overlay')) continue;
            const tx = b.innerText.toLowerCase();
            if (tx.includes('đọc sách') || tx.includes('đọc trực tuyến') || tx.includes('xem tài liệu')) return b;
        }
        return document.querySelector('button.css-5frqec') ||
            document.querySelector('button svg path[d*="M27.9197 15.8471"]')?.closest('button');
    };

    const findLikeButton = () => {
        const ariaSelectors = [
            'button[aria-label*="Thích" i]',
            'button[aria-label*="thích" i]',
            'button[aria-label*="Yêu thích" i]',
            'button[aria-label*="Like" i]',
            'button[title*="Thích" i]',
            'button[title*="thích" i]',
            'button[title*="Yêu thích" i]',
            'button[title*="Like" i]',
            '[role="button"][aria-label*="Thích" i]',
            '[role="button"][title*="Thích" i]'
        ];

        const readBtn = findReadButton();
        if (readBtn) {
            const containers = [readBtn.parentElement, readBtn.parentElement?.parentElement, readBtn.parentElement?.parentElement?.parentElement].filter(Boolean);
            for (const container of containers) {
                for (const sel of ariaSelectors) {
                    const b = container.querySelector(sel);
                    if (b && b !== readBtn && !b.closest('#tr-panel')) return b;
                }
                const svgs = Array.from(container.querySelectorAll('svg'));
                for (const svg of svgs) {
                    if (svg.closest('#tr-panel') || svg.closest(readBtn.tagName)) continue;
                    const testId = (svg.getAttribute('data-testid') || '').toLowerCase();
                    const svgHtml = svg.innerHTML.toLowerCase();
                    const isHeart = testId.includes('favorite') || testId.includes('heart') ||
                        svgHtml.includes('m12 21') || svgHtml.includes('m7.9996') ||
                        svgHtml.includes('16.5 3') || svgHtml.includes('4.5 2.09') ||
                        svgHtml.includes('21.35') || svgHtml.includes('m12 4.248');
                    if (isHeart) {
                        const btn = svg.closest('button, [role="button"], a, .MuiIconButton-root') || svg.parentElement;
                        if (btn && btn !== readBtn && !btn.closest('#tr-panel')) return btn;
                    }
                }
            }
        }

        const heartPaths = [
            'M7.9996 14.5879',
            'M12 21.35',
            'M12 21',
            '16.5 3',
            '4.5 2.09',
            '21.35l-1.45',
            'M12 4.248'
        ];
        for (const p of heartPaths) {
            const svgPaths = document.querySelectorAll(`svg path[d*="${p}"]`);
            for (const sp of svgPaths) {
                const btn = sp.closest('button, [role="button"], a, .MuiIconButton-root') || sp.closest('svg')?.parentElement;
                if (btn && !btn.closest('#tr-panel') && !btn.closest('header') && !btn.closest('nav')) return btn;
            }
        }

        const testIdSvgs = document.querySelectorAll('svg[data-testid*="Favorite"], svg[data-testid*="Heart"]');
        for (const svg of testIdSvgs) {
            const btn = svg.closest('button, [role="button"], a, .MuiIconButton-root') || svg.parentElement;
            if (btn && !btn.closest('#tr-panel') && !btn.closest('header') && !btn.closest('nav')) return btn;
        }

        for (const sel of ariaSelectors) {
            const btns = document.querySelectorAll(sel);
            for (const b of btns) {
                if (!b.closest('#tr-panel') && !b.closest('.tr-summary-box') && !b.closest('header') && !b.closest('nav')) return b;
            }
        }

        return null;
    };

    const isColorRedish = (colorStr) => {
        if (!colorStr) return false;
        const s = colorStr.toLowerCase();
        if (s.includes('red') || s.includes('#f44336') || s.includes('#d32f2f') || s.includes('#e91e63') || s.includes('#e11d48') || s.includes('#ef4444') || s.includes('#ff1744')) return true;
        const m = s.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
        if (m) {
            const r = parseInt(m[1], 10);
            const g = parseInt(m[2], 10);
            const b = parseInt(m[3], 10);
            if (r > 140 && r > g * 1.3 && r > b * 1.3) return true;
        }
        return false;
    };

    const isBookLiked = () => {
        try {
            const likeBtn = findLikeButton();
            if (!likeBtn) return false;

            const aria = (likeBtn.getAttribute('aria-label') || likeBtn.title || '').toLowerCase();
            if (aria.includes('bỏ thích') || aria.includes('bỏ yêu thích') || aria.includes('đã thích') || aria.includes('đã yêu thích') || aria.includes('unlike')) {
                return true;
            }

            if (likeBtn.getAttribute('aria-pressed') === 'true') return true;

            if (likeBtn.classList.contains('MuiIconButton-colorError') || likeBtn.classList.contains('Mui-selected') || likeBtn.classList.contains('liked') || likeBtn.classList.contains('active')) {
                return true;
            }

            const btnColor = window.getComputedStyle(likeBtn).color;
            if (isColorRedish(btnColor)) return true;

            const svg = likeBtn.querySelector('svg');
            if (svg) {
                const svgColor = window.getComputedStyle(svg).color;
                if (isColorRedish(svgColor)) return true;
                const svgFill = window.getComputedStyle(svg).fill;
                if (isColorRedish(svgFill)) return true;
                const attrFill = svg.getAttribute('fill') || svg.style.fill;
                if (isColorRedish(attrFill)) return true;
                const path = svg.querySelector('path');
                if (path) {
                    const pathFill = window.getComputedStyle(path).fill || path.getAttribute('fill');
                    if (isColorRedish(pathFill)) return true;
                }
            }

            return false;
        } catch (e) {
            return false;
        }
    };

    const autoLikeBook = async (updateStatus) => {
        try {
            let likeBtn = findLikeButton();
            for (let retry = 0; retry < 3 && !likeBtn; retry++) {
                await sleep(isFastMode() ? 100 : 350);
                likeBtn = findLikeButton();
            }

            if (isBookLiked()) {
                if (updateStatus) updateStatus(_d('IFPDoWNoIMSRw6MgdGjhuqMgdGltIHPhurVuIOKAlCBnaeG7ryBuZ3V5w6puIQ=='), 'info');
                return true;
            }

            if (!likeBtn) {
                if (updateStatus) updateStatus(_d('IEtow7RuZyB0w6xtIHRo4bqleSBuw7p0IHRpbSB0csOqbiB0cmFuZyE='), 'warn');
                return false;
            }

            const aria = (likeBtn.getAttribute('aria-label') || likeBtn.title || '').toLowerCase();
            if (aria.includes('bỏ thích') || aria.includes('bỏ yêu thích') || aria.includes('đã thích') || aria.includes('unlike')) {
                if (updateStatus) updateStatus(_d('IFPDoWNoIMSRw6MgdGjDrWNoIHLhu5NpIOKAlCBnaeG7ryBuZ3V5w6puIQ=='), 'info');
                return true;
            }

            if (updateStatus) updateStatus(_d('IMSQYW5nIHRo4bqjIHRpbSBjaG8gc8OhY2guLi4='), 'info');

            likeBtn.focus();
            const events = ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'];
            events.forEach(evName => {
                try {
                    likeBtn.dispatchEvent(new MouseEvent(evName, { bubbles: true, cancelable: true, view: window }));
                } catch (err) { }
            });
            try { likeBtn.click(); } catch (err) { }

            const innerSvg = likeBtn.querySelector('svg');
            if (innerSvg) {
                try { innerSvg.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); } catch (err) { }
            }

            await sleep(isFastMode() ? 300 : 900);

            const liked = isBookLiked();
            if (updateStatus) {
                if (liked) updateStatus(_d('IMSQw6MgdGjhuqMgdGltIHRow6BuaCBjw7RuZyE='), 'success');
                else updateStatus(_d('IMSQw6MgZ+G7rWkgbOG7h25oIGLhuqVtIHRpbSDigJQgdGnhur9wIHThu6VjIHjhu60gbMO9Li4u'), 'info');
            }
            return true;
        } catch (e) {
            return false;
        }
    };

    const isBookFullyReadAndCompleted = () => {
        const cond1 = isBookSlider100();
        const cond2 = isBookLiked();
        return cond1 && cond2;
    };

    const normalizeBookText = (text) => String(text || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();

    const getElementDistance = (first, second) => {
        const ancestors = new Map();
        let depth = 0;
        for (let node = first; node; node = node.parentElement) ancestors.set(node, depth++);
        depth = 0;
        for (let node = second; node; node = node.parentElement, depth++) {
            if (ancestors.has(node)) return ancestors.get(node) + depth;
        }
        return Infinity;
    };

    const getBookViewCount = (bookTitle) => {
        const candidates = Array.from(document.querySelectorAll('h6, p, span, [class*="MuiTypography"], [class*="view"], [class*="count"]'));
        const seenElements = new Set();
        const countElements = [];

        for (const element of candidates) {
            if (element.closest('#tr-panel') || element.closest('#tr-view-scrape-overlay') || element.closest('[class*="tr-"]')) continue;
            if (seenElements.has(element)) continue;

            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            if (style.display === 'none' || style.visibility === 'hidden' || rect.width <= 0 || rect.height <= 0) continue;

            const rawText = (element.innerText || element.textContent || '').replace(/\u00A0/g, ' ').trim();
            if (!rawText || rawText.length > 80) continue;

            const match = rawText.match(/(?:lượt\s*(?:xem|đọc)\s*[:：]?\s*([\d.,]+)\s*([kmb])?)|(?:([\d.,]+)\s*([kmb])?\s*lượt\s*(?:xem|đọc))/i);
            if (!match) continue;

            const numStr = match[1] || match[3];
            const suffix = (match[2] || match[4] || '').toLowerCase();
            const multiplier = suffix === 'k' ? 1000 : suffix === 'm' ? 1000000 : suffix === 'b' ? 1000000000 : 1;
            const parsed = suffix
                ? Number(numStr.replace(',', '.'))
                : Number(numStr.replace(/\D/g, ''));
            const count = Number.isFinite(parsed) ? Math.round(parsed * multiplier) : NaN;

            if (Number.isFinite(count)) {
                seenElements.add(element);
                countElements.push({ element, count });
            }
        }

        if (countElements.length === 0) return null;

        if (countElements.length === 1) return countElements[0].count;

        const readBtn = findReadButton();
        if (!bookTitle) {
            if (readBtn) {
                const nearRead = countElements.map(item => ({ ...item, dist: getElementDistance(readBtn, item.element) }))
                    .sort((a, b) => a.dist - b.dist);
                return nearRead[0].count;
            }
            return countElements[0].count;
        }

        const targetTitle = normalizeBookText(bookTitle);
        const titleElements = Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, [class*="title"]'));
        let titleElement = null;
        let bestTitleScore = 0;
        for (const element of titleElements) {
            if (element.closest('#tr-panel') || element.closest('#tr-view-scrape-overlay') || element.closest('[class*="tr-"]')) continue;
            const text = normalizeBookText(element.innerText || element.textContent);
            if (!text) continue;
            let score = 0;
            if (text === targetTitle) score = 100000 + text.length;
            else if (text.includes(targetTitle)) score = 90000 + targetTitle.length;
            else if (targetTitle.includes(text) && text.length >= 12) score = 80000 + text.length;
            if (score > bestTitleScore) {
                titleElement = element;
                bestTitleScore = score;
            }
        }

        if (titleElement) {
            const ranked = countElements.map((item) => ({ ...item, distance: getElementDistance(titleElement, item.element) }))
                .sort((a, b) => a.distance - b.distance);
            if (Number.isFinite(ranked[0]?.distance)) {
                if (ranked[0].distance !== ranked[1]?.distance || ranked[0].count === ranked[1]?.count) {
                    return ranked[0].count;
                }
            }
        }

        if (readBtn) {
            const nearRead = countElements.map(item => ({ ...item, dist: getElementDistance(readBtn, item.element) }))
                .sort((a, b) => a.dist - b.dist);
            return nearRead[0].count;
        }

        return countElements[0].count;
    };

    const waitForStableViewCount = async (bookTitle, maxWait = 8000, stableFor = 750, onSample = null) => {
        const start = Date.now();
        let lastCount = null;
        let changedAt = 0;
        while (Date.now() - start < maxWait) {
            if (isStopped()) return null;
            const currentCount = getBookViewCount(bookTitle);
            if (currentCount !== lastCount) {
                lastCount = currentCount;
                changedAt = Date.now();
                if (typeof onSample === 'function') onSample(currentCount);
            } else if (currentCount !== null && Date.now() - changedAt >= stableFor) {
                return currentCount;
            }
            await sleep(200);
        }
        return null;
    };

    const waitForElement = (selector, timeout = 10000) => new Promise((resolve) => {
        if (document.querySelector(selector)) return resolve(document.querySelector(selector));
        const obs = new MutationObserver(() => {
            if (document.querySelector(selector)) { obs.disconnect(); resolve(document.querySelector(selector)); }
        });
        obs.observe(document.body, { childList: true, subtree: true });
        setTimeout(() => { obs.disconnect(); resolve(null); }, timeout);
    });

    const waitForPdfLoad = async () => {
        const timeout = getPdfTimeout();
        const start = Date.now();
        const el = await new Promise((resolve) => {
            if (document.querySelector('label[for="df_book_page_number"]')) return resolve(document.querySelector('label[for="df_book_page_number"]'));
            if (hasFileAccessError()) return resolve(null);
            let settled = false;
            let timeoutId;
            let skipPoll;
            const finish = (value) => {
                if (settled) return;
                settled = true;
                obs.disconnect();
                clearTimeout(timeoutId);
                clearInterval(skipPoll);
                resolve(value);
            };
            const obs = new MutationObserver(() => {
                if (document.querySelector('label[for="df_book_page_number"]')) {
                    finish(document.querySelector('label[for="df_book_page_number"]'));
                } else if (hasFileAccessError()) {
                    finish(null);
                }
            });
            obs.observe(document.body, { childList: true, subtree: true });
            timeoutId = setTimeout(() => finish(null), timeout);
            skipPoll = setInterval(() => {
                if (isStopped()) finish(null);
            }, 100);
        });
        if (el) recordPdfTime(Date.now() - start);
        return el;
    };

    const waitForDocumentReaderLoad = async (updateStatus) => {
        const timeout = getPdfTimeout();
        const start = Date.now();
        while (Date.now() - start < timeout) {
            if (isStopped()) return null;
            if (wordReaderReadyAt) return { type: 'word', element: wordReaderReadyFrame };
            const pdfLabel = document.querySelector('label[for="df_book_page_number"]');
            if (pdfLabel) {
                recordPdfTime(Date.now() - start);
                return { type: 'pdf', element: pdfLabel };
            }

            const wordFrame = Array.from(document.querySelectorAll('iframe')).find(isOfficeWordFrame);
            if (wordFrame) {
                updateStatus(_d('IMSQYW5nIGNo4budIHTDoGkgbGnhu4d1IFdvcmQgdOG6o2kgeG9uZyAoU2hhcmVkQXV0aEZyYW1lKS4uLg=='), 'loading');
                while (Date.now() - start < timeout) {
                    if (isStopped()) return null;
                    if (wordReaderReadyAt) return { type: 'word', element: wordFrame };
                    const currentFrame = Array.from(document.querySelectorAll('iframe')).find(isOfficeWordFrame);
                    if (!currentFrame) break;
                    if (wordReaderReadyFrame === currentFrame) return { type: 'word', element: currentFrame };
                    await sleep(200);
                }
                return null;
            }
            if (hasFileAccessError()) return null;
            await sleep(100);
        }
        return null;
    };

    const waitForPageLoad = (targetPage, maxWait = 4000) => new Promise((resolve) => {
        const interval = isPerfMode() ? 50 : 100; let elapsed = 0;
        const text = targetPage + '/' + targetPage;
        const t = setInterval(() => {
            elapsed += interval;
            if (isStopped()) { clearInterval(t); return resolve(false); }
            if (hasFileAccessError()) { clearInterval(t); return resolve(false); }
            const label = document.querySelector('label[for="df_book_page_number"]');
            if (label && label.innerText.includes(text)) {
                clearInterval(t); setTimeout(() => resolve(true), isPerfMode() ? 50 : 150);
            } else if (elapsed >= maxWait) { clearInterval(t); resolve(false); }
        }, interval);
    });

    const findResumePopup = () => {
        const dialogs = Array.from(document.querySelectorAll('div[role="dialog"]'));
        for (const dlg of dialogs) {
            const text = (dlg.innerText || '').toLowerCase();
            if (!text.includes('xem tiếp') && !text.includes('đọc tiếp')) continue;
            const btns = Array.from(dlg.querySelectorAll('button'));
            const resetBtn = btns.find(b => b.innerText.trim() === 'Xem từ đầu');
            const contBtn = btns.find(b => {
                const t = b.innerText.trim().toLowerCase();
                return t === 'xem tiếp' || t === 'đọc tiếp' || t === 'tiếp tục' || t.includes('xem tiếp') || t.includes('đọc tiếp');
            });
            if (resetBtn || contBtn) return { dialog: dlg, resetBtn, contBtn };
        }
        const allBtns = Array.from(document.querySelectorAll('button'));
        const resetBtn = allBtns.find(b => b.innerText.trim() === 'Xem từ đầu');
        const contBtn = allBtns.find(b => { const t = b.innerText.trim().toLowerCase(); return t === 'xem tiếp' || t === 'đọc tiếp'; });
        if (resetBtn && contBtn) return { dialog: null, resetBtn, contBtn };
        return null;
    };
    const hasResumePopup = () => findResumePopup() !== null;
    const clickResumeFromStart = async () => {
        const popup = findResumePopup();
        if (!popup) return false;
        if (popup.resetBtn) { popup.resetBtn.click(); await sleep(200); return true; }
        if (popup.contBtn) { popup.contBtn.click(); await sleep(200); return true; }
        return false;
    };

    const handleLoginPopupIfNeeded = async (updateStatus) => {
        let elapsed = 0;
        while (elapsed < 2000) {
            if (isStopped()) return false;
            const dialog = document.querySelector('div[role="dialog"]');
            if (dialog && dialog.innerText.includes('Bạn cần đăng nhập trước khi đọc sách')) {
                updateStatus(_d('IFPDoWNoIHnDqnUgY+G6p3UgxJHEg25nIG5o4bqtcCDigJQgdOG7sSDEkcOzbmcgcG9wdXAuLi4='), 'warn');
                showErrorAlert('Sách yêu cầu đăng nhập trước khi đọc sách!');
                const btns = Array.from(dialog.querySelectorAll('button'));
                let close = btns.find(b => b.innerText.trim().toLowerCase() === 'đóng');
                if (!close) close = dialog.querySelector('button.MuiIconButton-root');
                if (close) { close.click(); await sleep(500); return true; }
            }
            await sleep(400); elapsed += 400;
        }
        return false;
    };

    const handleTeacherBookPopupIfNeeded = async (updateStatus) => {
        let elapsed = 0;
        while (elapsed < 2500) {
            if (isStopped()) return false;
            const dialog = document.querySelector('div[role="dialog"]');
            if (dialog) {
                const text = (dialog.innerText || '').toLowerCase();
                if (text.includes('tài khoản giáo viên') || text.includes('giáo viên để đọc sách')) {
                    if (updateStatus) updateStatus(_d('IFPDoWNoIGdpw6FvIHZpw6puIOKAlCDEkMOzbmcgcG9wdXAgJiBjaOG7iSB0aOG6oyB0aW0uLi4='), 'info');
                    const btns = Array.from(dialog.querySelectorAll('button'));
                    let close = btns.find(b => b.innerText.trim().toLowerCase() === 'đóng');
                    if (!close) close = dialog.querySelector('button.MuiIconButton-root');
                    if (close) { close.click(); await sleep(500); return true; }
                }
            }
            await sleep(300); elapsed += 300;
        }
        return false;
    };

    const scanSingleBook = async () => {
        await sleep(200);
        if (hasFileAccessError()) return { status: 'error', progress: null, reason: 'file_access_error' };

        const isFullyDone = isBookFullyReadAndCompleted();
        const slider100 = isBookSlider100();
        const liked = isBookLiked();
        const progress = getBookProgress();

        if (isFullyDone) {
            return { status: 'done', progress: 100, reason: 'slider 100% & đã tim ' };
        }
        if (slider100 && !liked) {
            await autoLikeBook();
            if (isBookLiked()) {
                return { status: 'done', progress: 100, reason: 'slider 100% & vừa bù tim ' };
            }
            return { status: 'partial', progress: 100, reason: 'slider 100% nhưng CHƯA TIM ' };
        }
        if (progress !== null && progress > 0) {
            return { status: 'partial', progress, reason: 'slider ' + progress + '%' + (liked ? ' (đã tim)' : ' (chưa tim)') };
        }

        const readBtn = findReadButton();
        if (!readBtn) return { status: 'not_started', progress, reason: 'no_read_btn' };

        if (readBtn.tagName === 'A') readBtn.removeAttribute('target');
        const originalOpen = window.open;
        window.open = () => null;
        try { readBtn.click(); } catch (e) { window.open = originalOpen; }

        let foundResume = false, foundLogin = false, foundTeacher = false, foundFileErr = false, elapsed = 0;
        while (elapsed < 2500) {
            if (isStopped()) break;
            if (hasFileAccessError()) { foundFileErr = true; break; }
            if (hasResumePopup()) { foundResume = true; break; }
            const dialog = document.querySelector('div[role="dialog"]');
            if (dialog) {
                const text = (dialog.innerText || '').toLowerCase();
                if (text.includes('tài khoản giáo viên') || text.includes('giáo viên để đọc sách')) {
                    foundTeacher = true;
                    const btns = Array.from(dialog.querySelectorAll('button'));
                    let close = btns.find(b => b.innerText.trim().toLowerCase() === 'đóng');
                    if (!close) close = dialog.querySelector('button.MuiIconButton-root');
                    if (close) close.click();
                    break;
                }
                if (text.includes('bạn cần đăng nhập trước khi đọc sách')) {
                    foundLogin = true;
                    const btns = Array.from(dialog.querySelectorAll('button'));
                    let close = btns.find(b => b.innerText.trim().toLowerCase() === 'đóng');
                    if (!close) close = dialog.querySelector('button.MuiIconButton-root');
                    if (close) close.click();
                    break;
                }
            }
            await sleep(250); elapsed += 250;
        }

        if (foundFileErr) { window.open = originalOpen; return { status: 'error', progress: null, reason: 'file_access_error' }; }
        if (foundTeacher) {
            await autoLikeBook();
            window.open = originalOpen;
            return { status: 'done', progress: 100, reason: 'sách giáo viên (chỉ tim)' };
        }
        if (foundResume) {
            const popup = findResumePopup();
            if (popup && popup.resetBtn) popup.resetBtn.click();
            await sleep(600); window.open = originalOpen;
            return { status: 'partial', progress, reason: 'resume_popup' };
        }
        if (foundLogin) { window.open = originalOpen; return { status: 'login_required', progress, reason: 'login' }; }
        const backBtn = Array.from(document.querySelectorAll('button.MuiIconButton-root')).find(b => b.innerHTML.includes('M0.93934 12.4393'));
        if (backBtn) backBtn.click();
        window.open = originalOpen;
        await sleep(500);
        return { status: 'not_started', progress, reason: 'no_popup' };
    };

    const showScrapingOverlay = () => {
        let overlay = document.getElementById('tr-scrape-overlay');
        if (overlay) overlay.remove();
        overlay = document.createElement('div');
        overlay.id = 'tr-scrape-overlay';
        overlay.innerHTML = '<div class="tr-scrape-message">Vui lòng chờ giây lát</div>';
        document.body.appendChild(overlay);
        return {
            close: () => {
                const el = document.getElementById('tr-scrape-overlay');
                if (el) el.remove();
            }
        };
    };

    const scrapeUntilExhausted = async (updateStatus, onCancelScrape) => {
        let lastCount = 0, noChangeStreak = 0, round = 0;
        const MAX_NO_CHANGE = 4, MAX_ROUNDS = 120;
        const scrapeOverlay = showScrapingOverlay();
        try {
            while (noChangeStreak < MAX_NO_CHANGE && round < MAX_ROUNDS) {
                if (isStopped()) {
                    if (typeof onCancelScrape === 'function') onCancelScrape(extractBooksFromPage());
                    return lastCount;
                }
                round++;
                const totalH = document.body.scrollHeight;
                for (let y = window.scrollY; y < totalH; y += 800) {
                    if (isStopped()) break;
                    window.scrollTo(0, y);
                    await sleep(120);
                }
                window.scrollTo(0, document.body.scrollHeight);
                await sleep(1400);

                const currentCount = extractBooksFromPage().length;
                if (currentCount > lastCount) {
                    noChangeStreak = 0;
                    updateStatus(_d('IMSQw6MgY8OgbyA=') + currentCount + _d('IHPDoWNoICh2w7JuZyA=') + round + ')...', 'info');
                } else {
                    noChangeStreak++;
                    updateStatus(_d('IEtp4buDbSB0cmEgY3Xhu5FpIHRyYW5nLi4uICg=') + noChangeStreak + '/' + MAX_NO_CHANGE + _d('KSDigJQg') + currentCount + _d('IHPDoWNo'), 'info');
                }
                lastCount = currentCount;

                const btns = Array.from(document.querySelectorAll('button, a'));
                let loadMore = null;
                for (const b of btns) {
                    const txt = (b.innerText || '').toLowerCase().trim();
                    if ((txt.includes('xem thêm') || txt.includes('tải thêm')) && !b.disabled && b.offsetParent !== null) { loadMore = b; break; }
                }
                if (loadMore) {
                    loadMore.click();
                    noChangeStreak = 0;
                    updateStatus(_d('IELhuqVtICJYZW0gdGjDqm0iIOKAlCDEkWFuZyB04bqjaS4uLg=='), 'info');
                    await sleep(2500);
                }
            }
            updateStatus(_d('IMSQYW5nIGPhuq1wIG5o4bqtdCBnw7NpIHTDoGkgbmd1ecOqbiBzw6FjaC4uLg=='), 'loading');
            await sleep(1500);
            const freshBooks = extractBooksFromPage();

            updateStatus(_d('IEPDoG8geG9uZzog') + lastCount + _d('IHPDoWNoICjEkcOjIGzGsHUgY2FjaGUp'), 'success');
            return lastCount;
        } finally {
            scrapeOverlay.close();
        }
    };

    const runScanStep = async (ui, updateStatus, updateProgress) => {
        const scanList = JSON.parse(localStorage.getItem(KEY_SCAN_LIST) || '[]');
        const scanIdx = parseInt(localStorage.getItem(KEY_SCAN_INDEX) || '0');
        const results = JSON.parse(localStorage.getItem(KEY_SCAN_RESULT) || '[]');
        if (!scanList[scanIdx]) { finishScan(scanList, results, ui, updateStatus); return; }

        if (hasFileAccessError()) {
            results.push({ title: scanList[scanIdx].title, url: scanList[scanIdx].url, progress: null, status: 'error', reason: 'file_access_error' });
            localStorage.setItem(KEY_SCAN_RESULT, JSON.stringify(results));
            const nextIdx = scanIdx + 1;
            localStorage.setItem(KEY_SCAN_INDEX, nextIdx.toString());
            updateProgress(ui, nextIdx, scanList.length);
            await sleep(400);
            if (nextIdx < scanList.length) window.location.href = scanList[nextIdx].url;
            else finishScan(scanList, results, ui, updateStatus);
            return;
        }

        updateProgress(ui, scanIdx, scanList.length);
        updateStatus(_d('IMSQYW5nIGtp4buDbSB0cmEg') + (scanIdx + 1) + '/' + scanList.length + _d('IOKAlCAi') + scanList[scanIdx].title.slice(0, 35) + _d('Li4uIg=='), 'loading');

        let result = { status: 'error', progress: null, reason: 'unknown' };
        try { result = await scanSingleBook(); }
        catch (e) {
            result = { status: 'error', progress: null, reason: e.message };
            showErrorAlert('Lỗi quét sách: ' + e.message);
        }

        results.push({ title: scanList[scanIdx].title, url: scanList[scanIdx].url, progress: result.progress, status: result.status, reason: result.reason });
        localStorage.setItem(KEY_SCAN_RESULT, JSON.stringify(results));

        const nextIdx = scanIdx + 1;
        localStorage.setItem(KEY_SCAN_INDEX, nextIdx.toString());
        updateProgress(ui, nextIdx, scanList.length);
        await sleep(600);
        if (nextIdx < scanList.length) window.location.href = scanList[nextIdx].url;
        else finishScan(scanList, results, ui, updateStatus);
    };

    const finishScan = (scanList, results, ui, updateStatus) => {
        localStorage.removeItem(KEY_SCAN_ACTIVE);
        localStorage.removeItem(KEY_SCAN_LIST);
        localStorage.removeItem(KEY_SCAN_INDEX);
        const done = results.filter(r => r.status === 'done').length;
        const partial = results.filter(r => r.status === 'partial').length;
        const notStarted = results.filter(r => r.status === 'not_started').length;
        const loginReq = results.filter(r => r.status === 'login_required').length;
        const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
        let autoAdded = 0;
        results.forEach(r => { if (r.status === 'done' && !history.includes(r.url)) { history.push(r.url); autoAdded++; } });
        localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
        setNormalMode(ui);
        updateStatus(_d('IFhvbmc6IA==') + done + _d('IMSR4bunIDIgxJBLIMK3IA==') + partial + _d('IGNoxrBhIMSR4bqhdCDCtyA=') + notStarted + _d('IGNoxrBhIMSR4buNYyDCtyA=') + loginReq + _d('IGPhuqduIMSQTg==') + (autoAdded > 0 ? _d('IOKAlCB04buxIGzGsHUg') + autoAdded + _d('IHPDoWNo') : ''), 'success');
        ui.etaWrap.classList.remove('tr-on');
        setTimeout(() => showScanReport(results), 700);
    };

    const buildReportHTML = (results) => {
        const done = results.filter(r => r.status === 'done');
        const partial = results.filter(r => r.status === 'partial');
        const notStarted = results.filter(r => r.status === 'not_started');
        const loginReq = results.filter(r => r.status === 'login_required');
        const errorList = results.filter(r => r.status === 'error');
        const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
        const renderRows = (list) => list.map(r => `<tr><td><a href="${esc(r.url)}" target="_blank" rel="noopener">${esc(r.title)}</a></td><td class="pct">${r.progress !== null ? r.progress + '%' : '--'}</td><td class="reason">${esc(r.reason || '')}</td></tr>`).join('');
        const section = (title, list) => list.length === 0 ? '' : `<div class="section"><div class="section-title">${title} <span class="count">${list.length}</span></div><table><thead><tr><th>Tên sách</th><th>%</th><th>Lý do</th></tr></thead><tbody>${renderRows(list)}</tbody></table></div>`;
        return `<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>Báo cáo — Auto Đọc Sách</title>
        <style>*{box-sizing:border-box}body{margin:0;padding:28px;background:#0F172A;color:#E2E8F0;font-family:'Roboto',-apple-system,'Segoe UI',sans-serif}
        h1{font-size:22px;margin:0 0 6px;color:#fff}.sub{color:#94A3B8;font-size:13px;margin-bottom:24px}
        .stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:24px}
        .stat{background:#1E293B;border:1px solid #334155;border-radius:14px;padding:14px 16px}
        .stat .num{font-size:26px;font-weight:800;line-height:1;margin-bottom:4px}
        .stat .lbl{font-size:11.5px;font-weight:600;letter-spacing:.4px;color:#94A3B8;text-transform:uppercase}
        .stat.done .num{color:#22C55E}.stat.partial .num{color:#F59E0B}.stat.not .num{color:#EF4444}.stat.login .num{color:#A855F7}.stat.total .num{color:#38BDF8}
        table{width:100%;border-collapse:collapse;background:#1E293B;border:1px solid #334155;border-radius:14px;overflow:hidden;margin-bottom:6px}
        th,td{padding:10px 14px;text-align:left;border-bottom:1px solid #334155;font-size:13px}
        th{background:#0F172A;font-weight:700;color:#94A3B8;text-transform:uppercase;font-size:11px;letter-spacing:.4px}
        tr:last-child td{border-bottom:none}tr:hover td{background:#293548}
        td a{color:#38BDF8;text-decoration:none}td a:hover{text-decoration:underline}
        td.pct{width:80px;font-variant-numeric:tabular-nums;font-weight:700}td.reason{width:180px;color:#94A3B8;font-size:11.5px}
        .section{margin-bottom:24px}.section-title{font-size:14px;font-weight:700;margin-bottom:10px;color:#E2E8F0;display:flex;align-items:center;gap:8px}
        .section-title .count{background:#334155;padding:2px 10px;border-radius:100px;font-size:11.5px;color:#CBD5E1}
        .btn{padding:9px 18px;border-radius:100px;background:#334155;color:#E2E8F0;border:none;cursor:pointer;font-family:inherit;font-size:12.5px;font-weight:600;margin-right:8px}.btn:hover{background:#475569}
        .note{padding:10px 14px;border-radius:10px;background:#1E293B;border:1px solid #334155;color:#94A3B8;font-size:12px;margin-bottom:20px}
        </style></head><body>
        <h1> Báo cáo kiểm tra đã đọc</h1>
        <div class="sub">${results.length} sách đã kiểm tra · Auto Đọc Sách V51</div>
        <div class="note"> <b>Quy chuẩn 2 điều kiện:</b> Phải thỏa mãn đồng thời <b>Slider 100%</b> VÀ <b>Đã Thả Tim</b> mới tính là "Đã đọc".</div>
        <div class="stats">
        <div class="stat total"><div class="num">${results.length}</div><div class="lbl">Tổng</div></div>
        <div class="stat done"><div class="num">${done.length}</div><div class="lbl"> Đã Đọc &amp; Tim</div></div>
        <div class="stat partial"><div class="num">${partial.length}</div><div class="lbl"> Chưa Đủ 2 ĐK</div></div>
        <div class="stat not"><div class="num">${notStarted.length}</div><div class="lbl"> Chưa đọc</div></div>
        <div class="stat login"><div class="num">${loginReq.length}</div><div class="lbl"> Cần ĐN</div></div>
        </div>
        <button class="btn" onclick="window.location.reload()"> Tải lại</button>
        ${section(' Đã Đọc 100% & Thả Tim', done)}
        ${section(' Chưa Đủ 2 ĐK (sẽ đọc/tim lại)', partial)}
        ${section(' Chưa đọc', notStarted)}
        ${section(' Cần đăng nhập', loginReq)}
        ${section(' Lỗi / File không truy cập được', errorList)}
        </body></html>`;
    };

    const showScanReport = (results) => {
        const html = buildReportHTML(results);
        try {
            const blob = new Blob([html], { type: 'text/html' });
            const url = URL.createObjectURL(blob);
            const win = window.open(url, '_blank');
            if (win) { setTimeout(() => URL.revokeObjectURL(url), 60000); return; }
        } catch (e) { }
        openReportOverlay(html);
    };
    const openReportOverlay = (html) => {
        const old = document.getElementById('tr-report-overlay');
        if (old) old.remove();
        const overlay = document.createElement('div');
        overlay.id = 'tr-report-overlay';
        const closeBtn = document.createElement('button');
        closeBtn.className = 'tr-report-close';
        closeBtn.innerHTML = '✕';
        closeBtn.onclick = () => overlay.remove();
        const iframe = document.createElement('iframe');
        iframe.srcdoc = html;
        iframe.className = 'tr-report-iframe';
        overlay.appendChild(closeBtn);
        overlay.appendChild(iframe);
        document.body.appendChild(overlay);
    };

    const showCountdown = (seconds) => new Promise((resolve) => {
        return resolve();
        const old = document.getElementById('tr-countdown-overlay');
        if (old) old.remove();
        const overlay = document.createElement('div');
        overlay.id = 'tr-countdown-overlay';
        const label = document.createElement('div');
        label.className = 'tr-cd-label';
        label.textContent = 'Auto sẽ tiếp tục sau';
        const num = document.createElement('div');
        num.className = 'tr-cd-num';
        num.textContent = seconds;
        overlay.appendChild(label); overlay.appendChild(num);
        document.body.appendChild(overlay);
        const start = Date.now();
        const totalMs = seconds * 1000;
        const interval = setInterval(() => {
            if (isStopped()) { clearInterval(interval); overlay.remove(); return resolve(); }
            const elapsed = Date.now() - start;
            const remain = Math.max(0, seconds - Math.floor(elapsed / 1000));
            if (num.textContent !== String(remain) && remain > 0) {
                num.textContent = remain;
                num.style.animation = 'none'; void num.offsetWidth;
                num.style.animation = 'tr-cd-num-pop .5s ease-out';
            }
            if (elapsed >= totalMs) {
                clearInterval(interval);
                overlay.classList.add('tr-cd-out');
                setTimeout(() => { overlay.remove(); resolve(); }, 250);
            }
        }, 100);
    });

    const waitForCoverLoaded = async (maxWait = 2000) => {
        const start = Date.now();
        await sleep(80);

        while (Date.now() - start < maxWait) {
            if (isStopped()) return false;
            if (hasFileAccessError()) return false;

            const input = document.getElementById('df_book_page_number');
            const hasPageNumber = input && input.value && input.value !== '0';

            if (isFastMode()) {
                const label = document.querySelector('label[for="df_book_page_number"]');
                const hasLabel = label && label.innerText && label.innerText.includes('/');
                if (hasPageNumber || hasLabel) {
                    await sleep(isPerfMode() ? 50 : 100);
                    return true;
                }
            }

            const loaders = Array.from(document.querySelectorAll('.df-loading, .df-page-loader, .df-ui-loader, .df-spinner, .loading-info'));
            const isStillLoading = loaders.some(el => {
                const style = window.getComputedStyle(el);
                return style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && (el.offsetWidth > 0 || el.offsetHeight > 0);
            });

            const canvas = document.querySelector('.df-container canvas, .df-book-stage canvas, canvas');
            const isCanvasReady = canvas && canvas.width > 50 && canvas.height > 50;

            if (isCanvasReady && !isStillLoading && hasPageNumber) {
                await sleep(isPerfMode() ? 100 : 200);
                return true;
            }

            await sleep(50);
        }
        await sleep(isPerfMode() ? 100 : 200);
        return false;
    };

    const waitForBookDetailsReady = async (maxWait = 3000) => {
        const start = Date.now();
        while (Date.now() - start < maxWait) {
            if (isStopped()) return false;
            if (findReadButton() || findLikeButton() || document.querySelector('label[for="df_book_page_number"]')) {
                await sleep(isFastMode() ? 80 : 150);
                return true;
            }
            await sleep(80);
        }
        return false;
    };

    const openBookIfNeeded = async (updateStatus) => {
        if (isStopped()) throw new Error('Đã dừng.');

        let readerLabel = document.querySelector('label[for="df_book_page_number"]');
        if (readerLabel) return 'pdf';
        if (wordReaderReadyAt || Array.from(document.querySelectorAll('iframe')).some(isOfficeWordFrame)) {
            const reader = await waitForDocumentReaderLoad(updateStatus);
            if (isStopped()) throw new Error('Đã dừng.');
            if (reader?.type === 'word') return 'word';
            throw new Error('Quá thời gian chờ tải tài liệu Word!');
        }

        if (isBookLiked()) {
            updateStatus(_d('IFPDoWNoIMSRw6MgdGjhuqMgdGltIHPhurVuIOKAlCBLSMOUTkcgYuG6pW0gdGltIGzhuqFpLCB2w6BvIMSR4buNYyBuZ2F5Li4u'), 'info');
        } else {
            updateStatus(_d('IENoxrBhIHRpbSDigJQgdOG7sSDEkeG7mW5nIHRo4bqjIHRpbS4uLg=='), 'info');
            await autoLikeBook(updateStatus);
            await sleep(isFastMode() ? 250 : 500);
        }

        updateStatus(_d('IMSQYW5nIHTDrG0gbsO6dCDEkOG7jWMgc8OhY2guLi4='), 'info');
        const readBtn = findReadButton();
        if (!readBtn) throw new Error('Không tìm thấy nút Đọc sách!');

        if (readBtn.tagName === 'A') readBtn.removeAttribute('target');
        window.open = (url) => { window.location.href = url; return window; };
        wordReaderReadyAt = 0;
        wordReaderReadyFrame = null;
        readBtn.click();

        let foundResume = false, foundLogin = false, foundTeacher = false, foundWord = false, waitMs = 0;
        while (waitMs < 2000) {
            if (isStopped()) throw new Error('Đã dừng.');
            if (hasResumePopup()) { foundResume = true; break; }
            const dialog = document.querySelector('div[role="dialog"]');
            if (dialog) {
                const text = (dialog.innerText || '').toLowerCase();
                if (text.includes('tài khoản giáo viên') || text.includes('giáo viên để đọc sách')) {
                    foundTeacher = true; break;
                }
                if (text.includes('bạn cần đăng nhập trước khi đọc sách')) {
                    foundLogin = true; break;
                }
            }
            if (Array.from(document.querySelectorAll('iframe')).some(isOfficeWordFrame)) { foundWord = true; break; }
            if (wordReaderReadyAt) { foundWord = true; break; }
            if (document.querySelector('label[for="df_book_page_number"]')) break;
            await sleep(100); waitMs += 100;
        }

        if (foundTeacher) {
            await handleTeacherBookPopupIfNeeded(updateStatus);
            if (!isBookLiked()) await autoLikeBook(updateStatus);
            throw new Error('TEACHER_BOOK');
        }

        if (foundLogin) {
            await handleLoginPopupIfNeeded(updateStatus);
            throw new Error('Sách yêu cầu đăng nhập, đã bỏ qua!');
        }
        if (foundWord) {
            const reader = await waitForDocumentReaderLoad(updateStatus);
            if (isStopped()) throw new Error('Đã dừng.');
            if (reader?.type === 'word') return 'word';
            throw new Error('Quá thời gian chờ tải tài liệu Word!');
        }
        if (foundResume) {
            updateStatus(_d('IE3hu58gbOG6oWkgdOG7qyDEkeG6p3UuLi4='), 'info');
            await clickResumeFromStart();
            await sleep(200);
            updateStatus(_d('IMSQYW5nIGNo4budIFBERiBt4bufLi4u'), 'loading');
            const reader = await waitForDocumentReaderLoad(updateStatus);
            if (isStopped()) throw new Error('Đã dừng.');
            if (reader?.type === 'word') return 'word';
            readerLabel = reader?.type === 'pdf' ? reader.element : null;
            if (!readerLabel) throw new Error('Quá thời gian chờ tải PDF sau popup!');
            return 'pdf';
        }

        if (isStopped()) throw new Error('Đã dừng.');
        updateStatus(_d('IMSQYW5nIGNo4budIFBERiB04bqjaSBraHVuZy4uLg=='), 'loading');
        const reader = await waitForDocumentReaderLoad(updateStatus);
        if (isStopped()) throw new Error('Đã dừng.');
        if (reader?.type === 'word') return 'word';
        readerLabel = reader?.type === 'pdf' ? reader.element : null;
        if (!readerLabel) throw new Error('Quá thời gian chờ tải PDF!');
        return 'pdf';
    };

    const performAutoReadLogic = async (updateStatus, opts = {}) => {
        const retryCount = opts._retryCount || 0;
        const MAX_VERIFY_RETRY = 3;
        let viewCountBefore = null;

        document.querySelector('.tr-panel')?.classList.add('tr-panel-expanded');

        if (isStopped()) throw new Error('Đã dừng.');

        if (!opts.forceRead && isBookFullyReadAndCompleted()) {
            updateStatus(_d('IFPDoWNoIMSRw6MgaG/DoG4gdOG6pXQg4oCUIGLhu48gcXVhIHThu6ljIHRow6wh'), 'success');
            throw new Error('ALREADY_READ');
        }

        updateStatus(_d('IENo4budIGdpYW8gZGnhu4duIHPDoWNoIHThuqNpIHhvbmcuLi4='), 'loading');
        await waitForBookDetailsReady();
        if (isStopped()) throw new Error('Đã dừng.');
        if (opts.verifyViewCount) {
            opts.onViewCountDebug?.({ title: opts.bookTitle, before: null, after: null, observed: null, status: 'Đang chờ lượt xem trước khi đọc...' });
            viewCountBefore = await waitForStableViewCount(opts.bookTitle, 6000, 750, (count) => {
                opts.onViewCountDebug?.({ observed: count, status: count === null ? 'Chưa nhận diện được số lượt xem' : 'Đang đọc số lượt xem trước...' });
            });
            opts.onViewCountDebug?.({
                before: viewCountBefore,
                observed: viewCountBefore,
                status: viewCountBefore === null ? 'Không đọc được lượt xem ban đầu' : 'Đã ghi nhận lượt xem ban đầu'
            });
            if (isStopped()) throw new Error('Đã dừng.');
        }

        updateStatus(_d('IEtp4buDbSB0cmEgxJFp4buBdSBraeG7h24gaG/DoG4gdOG6pXQuLi4='), 'info');

        if (!opts.forceRead) {
            const isCompleted = isBookFullyReadAndCompleted();
            if (isCompleted) {
                updateStatus(_d('IMSQw6MgaG/DoG4gdGjDoG5oIHRyxrDhu5tjIMSRw7Mg4oCUIGLhu48gcXVhIG5nYXk='), 'success');
                throw new Error('ALREADY_READ');
            }

            if (isBookSlider100() && !isBookLiked()) {
                updateStatus(_d('IFPDoWNoIMSRw6MgMTAwJSDigJQgdOG7sSDEkeG7mW5nIGLhuqVtIHRpbSBiw7kgxJHhu4MgaG/DoG4gdOG6pXQuLi4='), 'warn');
                await autoLikeBook(updateStatus);
                if (isBookLiked() || isFastMode()) {
                    updateStatus(_d('IMSQw6MgxJHhu6cgU2xpZGVyIDEwMCUgJiDEkMOjIFRpbSDigJQgaG/DoG4gdOG6pXQgc8OhY2gh'), 'success');
                    await sleep(isFastMode() ? 250 : 500);
                    return;
                }
            } else if (getBookProgress() > 0) {
                updateStatus(_d('IFPDoWNoIMSR4buNYyBk4bufIOKAlCDEkeG7jWMgbOG6oWkgdOG7qyDEkeG6p3U='), 'warn');
            }
        }

        const readerType = await openBookIfNeeded(updateStatus);
        if (isStopped()) throw new Error('Đã dừng.');

        if (readerType === 'word') {
            updateStatus(_d('VMOgaSBsaeG7h3UgV29yZCDEkcOjIHThuqNpIHhvbmc7IG5o4bqtbiBkaeG7h24gcXVhIFNoYXJlZEF1dGhGcmFtZS4='), 'success');
            const allBtns = document.querySelectorAll('button.MuiIconButton-root');
            let closeBtn = null;
            for (const button of allBtns) {
                if (button.innerHTML.includes('M0.93934 12.4393')) { closeBtn = button; break; }
            }
            if (!closeBtn) {
                closeBtn = document.querySelector('button[aria-label*="Trở về" i], button[aria-label*="Quay lại" i], button[aria-label*="Đóng" i], button[title*="Trở về" i], button[title*="Quay lại" i], button[title*="Đóng" i]');
            }
            if (closeBtn) {
                closeBtn.dispatchEvent(new MouseEvent('click', { view: window, bubbles: true, cancelable: true }));
                await sleep(isPerfMode() ? 150 : 300);
            }
            if (!isBookLiked()) await autoLikeBook(updateStatus);
            if (isStopped()) throw new Error('Đã dừng.');
            updateStatus(_d('xJDDoyBuaOG6rW4gZGnhu4duIHTDoGkgbGnhu4d1IFdvcmQgdOG6o2kgeG9uZyBxdWEgU2hhcmVkQXV0aEZyYW1lLg=='), 'success');
            if (opts.verifyViewCount) {
                updateStatus(_d('xJBhbmcga2nhu4NtIHRyYSBsxrDhu6N0IHhlbSBzYXUga2hpIMSR4buNYy4uLg=='), 'loading');
                opts.onViewCountDebug?.({ after: null, status: 'Đang chờ lượt xem sau khi đọc...' });
                const maxWaitWord = (viewCountBefore && viewCountBefore >= 1000) ? 2500 : 12000;
                const viewCountAfter = await waitForStableViewCount(opts.bookTitle, maxWaitWord, 750, (count) => {
                    opts.onViewCountDebug?.({ observed: count, status: count === null ? 'Chưa nhận diện được lượt xem sau' : 'Đang đọc số lượt xem sau...' });
                });
                opts.onViewCountDebug?.({
                    after: viewCountAfter,
                    observed: viewCountAfter,
                    status: viewCountAfter === null ? 'Không đọc được lượt xem sau' : 'Đã ghi nhận lượt xem sau'
                });
                if (isStopped()) throw new Error('Đã dừng.');
                return { viewCountBefore, viewCountAfter };
            }
            return;
        }

        let label = document.querySelector('label[for="df_book_page_number"]');
        let totalPages = label?.innerText.split('/')[1]?.trim();
        if (!totalPages || totalPages === '0' || totalPages === 'undefined') {
            await sleep(isPerfMode() ? 100 : 250);
            label = document.querySelector('label[for="df_book_page_number"]');
            totalPages = label?.innerText.split('/')[1]?.trim();
        }
        if (!totalPages) throw new Error('Lỗi đếm số trang!');

        updateStatus(_d('IMSQYW5nIGNo4budIHThuqNpIHhvbmcgdHJhbmcgYsOsYS4uLg=='), 'loading');
        await waitForCoverLoaded(isPerfMode() ? 1000 : 1200);
        if (isStopped()) throw new Error('Đã dừng.');

        updateStatus(_d('IMSQw6MgdOG6o2kgeG9uZyBiw6xhIOKAlCDEkWFuZyDEkcOzbmcgc8OhY2guLi4='), 'success');
        const allBtns = document.querySelectorAll('button.MuiIconButton-root');
        let backBtn = null;
        for (const b of allBtns) { if (b.innerHTML.includes('M0.93934 12.4393')) { backBtn = b; break; } }
        if (backBtn) {
            backBtn.dispatchEvent(new MouseEvent('click', { view: window, bubbles: true, cancelable: true }));
            await sleep(isPerfMode() ? 80 : 150);
        } else {
            const fallbackBack = document.querySelector('button[aria-label*="Trở về" i], button[aria-label*="Quay lại" i], button[title*="Trở về" i], button[title*="Quay lại" i]');
            if (fallbackBack) {
                fallbackBack.dispatchEvent(new MouseEvent('click', { view: window, bubbles: true, cancelable: true }));
                await sleep(isPerfMode() ? 80 : 150);
            }
        }
        if (!isBookLiked()) {
            await autoLikeBook(updateStatus);
        }

        updateStatus(_d('IMSQw6MgdOG6o2kgYsOsYSAmIMSRw7NuZyBzw6FjaCB0aMOgbmggY8O0bmcg4oCUIEhvw6BuIHThuqV0IQ=='), 'success');
        if (opts.verifyViewCount) {
            updateStatus(_d('IMSQYW5nIGtp4buDbSB0cmEgbMaw4bujdCB4ZW0gc2F1IGtoaSDEkeG7jWMuLi4='), 'loading');
            opts.onViewCountDebug?.({ after: null, status: 'Đang chờ lượt xem sau khi đọc...' });
            const maxWaitPdf = (viewCountBefore && viewCountBefore >= 1000) ? 2500 : 10000;
            const viewCountAfter = await waitForStableViewCount(opts.bookTitle, maxWaitPdf, 750, (count) => {
                opts.onViewCountDebug?.({ observed: count, status: count === null ? 'Chưa nhận diện được lượt xem sau' : 'Đang đọc số lượt xem sau...' });
            });
            opts.onViewCountDebug?.({
                after: viewCountAfter,
                observed: viewCountAfter,
                status: viewCountAfter === null ? 'Không đọc được lượt xem sau' : 'Đã ghi nhận lượt xem sau'
            });
            if (isStopped()) throw new Error('Đã dừng.');
            return { viewCountBefore, viewCountAfter };
        }
        return;
    };

    const extractBooksFromPage = () => {
        const list = [];
        const seen = new Set();
        const links = document.querySelectorAll('a');
        const exclude = ['sach-dien-tu', 'sach-noi', 'bai-giang-dien-tu', 'album', 'video', 'moi-nhat', 'xem-nhieu'];
        const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
        let historyUpdated = false;

        links.forEach(a => {
            const url = a.href;
            if (!url) return;
            if (isListPage(url)) return;
            if (url.includes('/chi-tiet/') || url.includes('/tai-lieu/')) {
                const isCat = exclude.some(p => url.endsWith('/' + p) || url.includes('/' + p + '?'));
                const isRoot = url.endsWith('/tai-lieu') || url.endsWith('/tai-lieu/');
                if (isCat || isRoot) return;
                if (seen.has(url)) return;
                const txt = a.innerText.replace(/\n/g, ' ').trim();
                if (txt && txt.length > 3) {
                    const card = a.closest('.MuiCard-root, .card, [class*="card"], [class*="item"]') || a.parentElement?.parentElement || a.parentElement;
                    let isCardRead = false;
                    if (card) {
                        const cardText = card.innerText.toLowerCase();
                        if (cardText.includes('100%') || cardText.includes('đã đọc') || cardText.includes('đã hoàn thành') || cardText.includes('hoàn tất')) {
                            isCardRead = true;
                        }
                        const tracks = card.querySelectorAll('.MuiSlider-track, .MuiLinearProgress-bar, [role="progressbar"]');
                        for (const tr of tracks) {
                            const w = tr.style.width || tr.getAttribute('aria-valuenow') || '';
                            if (w === '100%' || w === '100' || w.startsWith('100')) { isCardRead = true; break; }
                        }
                    }

                    if (isCardRead && !history.includes(url)) {
                        history.push(url);
                        historyUpdated = true;
                    }

                    list.push({ uid: generateBookUID(url, txt), title: txt, url, isRead: isCardRead || history.includes(url) });
                    seen.add(url);
                }
            }
        });

        if (historyUpdated) {
            localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
        }

        return list;
    };

    const showBookPicker = (rawBooks, customTitle) => new Promise((resolve) => {
        const old = document.getElementById('tr-picker-overlay');
        if (old) old.remove();
        const currentLimit = getLimit();
        const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
        const overlay = document.createElement('div');
        overlay.id = 'tr-picker-overlay';
        const box = document.createElement('div');
        box.className = 'tr-picker-box';

        box.innerHTML = `
                    <div class="tr-picker-header">
                        <div>
                            <div class="tr-picker-title">${customTitle || ' Chọn sách cần đọc'}</div>
                            <div class="tr-picker-sub">${rawBooks.length} sách · ${history.length} sách đã hoàn tất (100% & tim)</div>
                        </div>
                    </div>
                    <div class="tr-picker-controls">
                        <div class="tr-picker-ctrl">
                            <label>Số sách cần đọc:</label>
                            <input type="number" id="tr-picker-limit" min="0" max="${rawBooks.length}" value="${currentLimit}" placeholder="0 = tất cả">
                            <span class="tr-picker-hint">0 = không giới hạn</span>
                        </div>
                        <div class="tr-picker-actions">
                            <button type="button" class="tr-picker-mini" id="tr-pick-select-all">Chọn tất cả</button>
                            <button type="button" class="tr-picker-mini" id="tr-pick-clear-all">Bỏ chọn tất cả</button>
                        </div>
                    </div>
                    <div class="tr-picker-list" id="tr-pick-list"></div>
                    <div class="tr-picker-footer">
                        <div class="tr-picker-count" id="tr-pick-count">0 đã chọn</div>
                        <div>
                            <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-pick-cancel">Huỷ</button>
                            <button class="tr-btn tr-btn-primary tr-btn-sm" id="tr-pick-ok">Bắt đầu đọc</button>
                        </div>
                    </div>`;

        const list = box.querySelector('#tr-pick-list');
        const filteredBooks = rawBooks.map((b, i) => ({ ...b, idx: i, read: history.includes(b.url) }));
        filteredBooks.forEach((b, i) => {
            const item = document.createElement('label');
            item.className = 'tr-picker-item' + (b.read ? ' tr-read' : '');
            item.innerHTML = `<input type="checkbox" data-idx="${i}" ${b.read ? '' : 'checked'}>
                        <div class="tr-picker-item-info">
                            <div class="tr-picker-item-title">${b.title.length > 70 ? b.title.slice(0, 70) + '...' : b.title}</div>
                            <div class="tr-picker-item-meta">${b.read ? ' Đã đọc 100% & Tim' : ' Chưa hoàn tất'}</div>
                        </div>`;
            list.appendChild(item);
        });

        overlay.appendChild(box);
        document.body.appendChild(overlay);

        const countEl = box.querySelector('#tr-pick-count');
        const updateCount = () => {
            const n = box.querySelectorAll('input[type="checkbox"]:checked').length;
            countEl.innerText = n + ' đã chọn';
        };
        box.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.addEventListener('change', updateCount));
        updateCount();
        box.querySelector('#tr-pick-select-all').onclick = () => {
            box.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = true; });
            updateCount();
        };
        box.querySelector('#tr-pick-clear-all').onclick = () => {
            box.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; });
            updateCount();
        };
        box.querySelector('#tr-pick-cancel').onclick = () => { overlay.remove(); resolve(null); };
        box.querySelector('#tr-pick-ok').onclick = () => {
            const selected = [];
            box.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
                const idx = parseInt(cb.dataset.idx, 10);
                selected.push(filteredBooks[idx]);
            });
            const limitVal = parseInt(box.querySelector('#tr-picker-limit').value, 10) || 0;
            const limit = Math.max(0, limitVal);
            setLimit(limit);
            overlay.remove();
            resolve({ selectedUrls: new Set(selected.map(s => s.url)), limit });
        };
    });

    const showConfirmDialog = (title, message, okText = 'Tiếp tục', cancelText = 'Huỷ') => new Promise((resolve) => {
        const old = document.getElementById('tr-confirm-overlay');
        if (old) old.remove();
        const overlay = document.createElement('div');
        overlay.id = 'tr-confirm-overlay';
        const box = document.createElement('div');
        box.className = 'tr-confirm-box';
        box.innerHTML = `
                    <div class="tr-confirm-header">
                        <div class="tr-confirm-icon"><svg viewBox="0 0 24 24" width="30" height="30" fill="#F59E0B"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg></div>
                        <div class="tr-confirm-title">${title}</div>
                    </div>
                    <div class="tr-confirm-body">${message}</div>
                    <div class="tr-confirm-footer">
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-confirm-cancel">${cancelText}</button>
                        <button class="tr-btn tr-btn-primary tr-btn-sm" id="tr-confirm-ok">${okText}</button>
                    </div>`;
        overlay.appendChild(box);
        document.body.appendChild(overlay);
        box.querySelector('#tr-confirm-cancel').addEventListener('click', () => { overlay.remove(); resolve(false); });
        box.querySelector('#tr-confirm-ok').addEventListener('click', () => { overlay.remove(); resolve(true); });
        overlay.addEventListener('click', (e) => { if (e.target === overlay) { overlay.remove(); resolve(false); } });
    });

    const showAutoModeDialog = () => new Promise((resolve) => {
        const old = document.getElementById('tr-mode-overlay');
        if (old) old.remove();
        const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
        const overlay = document.createElement('div');
        overlay.id = 'tr-mode-overlay';
        const box = document.createElement('div');
        box.className = 'tr-mode-box';
        box.innerHTML = `
                    <div class="tr-mode-header">
                        <div class="tr-mode-icon">
                            <svg viewBox="0 0 24 24" width="28" height="28" fill="currentColor">
                                <path d="M11 21h-1l1-7H7.5c-.58 0-.57-.32-.38-.66.19-.34.05-.08.07-.12C8.48 10.94 10.42 7.54 13 3h1l-1 7h3.5c.49 0 .56.33.47.51l-.07.15C12.9 17.55 11 21 11 21z"/>
                            </svg>
                        </div>
                        <div class="tr-mode-title">Chọn chế độ Auto</div>
                        <div class="tr-mode-sub">Chọn cách bạn muốn đọc sách tự động</div>
                    </div>
                    <div class="tr-mode-options">
                        <div class="tr-mode-card" data-mode="1">
                            <div class="tr-mode-card-icon tr-mode-blue">
                                <svg viewBox="0 0 24 24" width="22" height="22" fill="currentColor">
                                    <path d="M21 5c-1.11-.35-2.33-.5-3.5-.5-1.95 0-4.05.4-5.5 1.5-1.45-1.1-3.55-1.5-5.5-1.5S2.45 4.9 1 6v14.65c0 .25.25.5.5.5.1 0 .15-.05.25-.05C3.1 20.45 5.05 20 6.5 20c1.95 0 4.05.4 5.5 1.5 1.35-.85 3.8-1.5 5.5-1.5 1.65 0 3.35.3 4.75 1.05.1.05.15.05.25.05.25 0 .5-.25.5-.5V6c-.6-.45-1.25-.75-2-1zm-1 13.5c-1.1-.35-2.3-.5-3.5-.5-1.7 0-4.15.65-5.5 1.5V8c1.35-.85 3.8-1.5 5.5-1.5 1.2 0 2.4.15 3.5.5v11.5z"/>
                                </svg>
                            </div>
                            <div class="tr-mode-card-body">
                                <div class="tr-mode-card-title-row">
                                    <span class="tr-mode-card-title">Chỉ sách chưa đọc</span>
                                    <span class="tr-mode-pill tr-pill-blue">Tự động</span>
                                </div>
                                <div class="tr-mode-card-desc">Tự tìm và đọc tất cả sách chưa hoàn tất. Quá trình này sẽ tốn thời gian vì vừa kiếm sách chưa đọc vừa đọc.</div>
                            </div>
                            <div class="tr-mode-card-arrow">
                                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                                    <path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/>
                                </svg>
                            </div>
                        </div>
                        <div class="tr-mode-card" data-mode="2">
                            <div class="tr-mode-card-icon tr-mode-green">
                                <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                                    <circle cx="12" cy="12" r="10"/>
                                    <circle cx="12" cy="12" r="6"/>
                                    <circle cx="12" cy="12" r="2"/>
                                </svg>
                            </div>
                            <div class="tr-mode-card-body">
                                <div class="tr-mode-card-title-row">
                                    <span class="tr-mode-card-title">Đọc theo số lượng</span>
                                    <span class="tr-mode-pill tr-pill-green">Mục tiêu</span>
                                </div>
                                <div class="tr-mode-card-desc">Nhập số cuốn cần đọc. Cuốn đã đọc hoặc sách giáo viên sẽ tự skip và không tính vào số lượng.</div>
                            </div>
                            <div class="tr-mode-card-arrow">
                                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                                    <path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/>
                                </svg>
                            </div>
                        </div>
                        <div class="tr-mode-card" data-mode="3">
                            <div class="tr-mode-card-icon tr-mode-purple">
                                <svg viewBox="0 0 24 24" width="22" height="22" fill="currentColor">
                                    <path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/>
                                </svg>
                            </div>
                            <div class="tr-mode-card-body">
                                <div class="tr-mode-card-title-row">
                                    <span class="tr-mode-card-title">Chọn sách từ danh sách</span>
                                    <span class="tr-mode-pill tr-pill-purple">Chỉ định</span>
                                </div>
                                <div class="tr-mode-card-desc">Hiện danh sách sách và tự chọn cuốn cần đọc. Cuốn đã đọc nếu chọn sẽ tự skip.</div>
                            </div>
                            <div class="tr-mode-card-arrow">
                                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                                    <path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/>
                                </svg>
                            </div>
                        </div>
                        <div class="tr-mode-card" data-mode="4">
                            <div class="tr-mode-card-icon tr-mode-cyan">
                                <svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M2 12s3.6-6 10-6 10 6 10 6-3.6 6-10 6S2 12 2 12z"/>
                                    <circle cx="12" cy="12" r="2.5"/>
                                </svg>
                            </div>
                            <div class="tr-mode-card-body">
                                <div class="tr-mode-card-title-row">
                                    <span class="tr-mode-card-title">Cào lượt đọc</span>
                                    <span class="tr-mode-pill tr-pill-cyan">Thống kê</span>
                                </div>
                                <div class="tr-mode-card-desc">Mở lại từ đầu các cuốn sách và kiểm tra lượt xem trước, sau khi đọc. Cuốn không tăng lượt sẽ có trong báo cáo.</div>
                            </div>
                            <div class="tr-mode-card-arrow">
                                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                                    <path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/>
                                </svg>
                            </div>
                        </div>
                    </div>
                    <div class="tr-mode-footer" style="justify-content: flex-end;">
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-mode-cancel">Huỷ</button>
                    </div>`;
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        box.querySelectorAll('.tr-mode-card').forEach(card => {
            card.addEventListener('click', async () => {
                const selectedMode = card.dataset.mode;
                if (selectedMode === '1') {
                    const ok = await showConfirmDialog(
                        'Xác nhận chế độ',
                        'Quá trình này sẽ rất tốn thời gian vì vừa kiếm sách chưa đọc mà vừa đọc. Bạn có chắc muốn tiếp tục không?',
                        'Tiếp tục',
                        'Huỷ'
                    );
                    if (!ok) return;
                }
                overlay.remove();
                resolve(selectedMode);
            });
        });
        box.querySelector('#tr-mode-cancel').addEventListener('click', () => { overlay.remove(); resolve(null); });
        overlay.addEventListener('click', (e) => { if (e.target === overlay) { overlay.remove(); resolve(null); } });
    });

    const showSummaryDialog = ({ successCount = 0, durationText = '--', totalTarget = 0 }) => new Promise((resolve) => {
        const old = document.getElementById('tr-summary-overlay');
        if (old) old.remove();
        const overlay = document.createElement('div');
        overlay.id = 'tr-summary-overlay';
        const box = document.createElement('div');
        box.className = 'tr-summary-box';

        box.innerHTML = `
                    <div class="tr-summary-header">
                        <div class="tr-summary-icon"><svg viewBox="0 0 24 24" width="42" height="42" fill="#16A34A"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/></svg></div>
                        <div class="tr-summary-title">Tổng kết phiên đọc sách</div>
                        <div class="tr-summary-sub">${totalTarget > 0 ? `Đã hoàn thành mục tiêu đọc ${successCount}/${totalTarget} cuốn` : `Đã hoàn tất phiên đọc sách`}</div>
                    </div>
                    <div class="tr-summary-stats">
                        <div class="tr-summary-stat-card stat-success">
                            <div class="tr-summary-stat-val">${successCount}</div>
                            <div class="tr-summary-stat-lbl">Đã đọc xong</div>
                        </div>
                        <div class="tr-summary-stat-card stat-time">
                            <div class="tr-summary-stat-val">${durationText}</div>
                            <div class="tr-summary-stat-lbl">Thời gian đọc</div>
                        </div>
                    </div>
                    <div class="tr-summary-footer">
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-summary-list-btn">
                            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>
                            <span>Về danh sách sách</span>
                        </button>
                        <button class="tr-btn tr-btn-primary tr-btn-sm" id="tr-summary-close-btn">Đóng</button>
                    </div>
                `;
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        box.querySelector('#tr-summary-list-btn').addEventListener('click', () => {
            overlay.remove();
            window.location.href = LIST_URL;
            resolve();
        });
        box.querySelector('#tr-summary-close-btn').addEventListener('click', () => {
            overlay.remove();
            resolve();
        });
    });

    const showNumberInputDialog = (maxBooks) => new Promise((resolve) => {
        const old = document.getElementById('tr-num-overlay');
        if (old) old.remove();
        const overlay = document.createElement('div');
        overlay.id = 'tr-num-overlay';
        const box = document.createElement('div');
        box.className = 'tr-num-box';
        box.innerHTML = `
                    <div class="tr-num-header">
                        <div class="tr-num-icon"><svg viewBox="0 0 24 24" width="36" height="36" fill="none" stroke="#0B57D0" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg></div>
                        <div class="tr-num-title">Nhập số cuốn cần đọc</div>
                        <div class="tr-num-sub">Có ${maxBooks} cuốn chưa đọc</div>
                    </div>
                    <div class="tr-num-input-wrap">
                        <input type="number" id="tr-num-input" min="1" max="${maxBooks}" value="5" placeholder="Nhập số...">
                        <div class="tr-num-hint">Tối đa ${maxBooks} cuốn chưa đọc</div>
                    </div>
                    <div class="tr-num-footer">
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-num-cancel">Huỷ</button>
                        <button class="tr-btn tr-btn-primary tr-btn-sm" id="tr-num-ok">Bắt đầu đọc</button>
                    </div>`;
        overlay.appendChild(box);
        document.body.appendChild(overlay);
        const input = box.querySelector('#tr-num-input');
        input.focus(); input.select();
        box.querySelector('#tr-num-cancel').addEventListener('click', () => { overlay.remove(); resolve(null); });
        box.querySelector('#tr-num-ok').addEventListener('click', () => {
            const val = parseInt(input.value, 10);
            if (!val || val < 1) { input.focus(); return; }
            overlay.remove(); resolve(Math.min(val, maxBooks));
        });
        input.addEventListener('keydown', (e) => { if (e.key === 'Enter') box.querySelector('#tr-num-ok').click(); });
        overlay.addEventListener('click', (e) => { if (e.target === overlay) { overlay.remove(); resolve(null); } });
    });

    const showLoopInputDialog = (bookCount) => new Promise((resolve) => {
        const old = document.getElementById('tr-loop-overlay');
        if (old) old.remove();
        const overlay = document.createElement('div');
        overlay.id = 'tr-loop-overlay';
        const box = document.createElement('div');
        box.className = 'tr-num-box';
        box.innerHTML = `
                    <div class="tr-num-header">
                        <div class="tr-num-icon"><svg viewBox="0 0 24 24" width="36" height="36" fill="none" stroke="#0B57D0" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg></div>
                        <div class="tr-num-title">Số vòng lặp cào lượt đọc</div>
                        <div class="tr-num-sub">Danh sách gồm <strong>${bookCount} cuốn sách</strong></div>
                    </div>
                    <div class="tr-num-input-wrap">
                        <input type="number" id="tr-loop-input" min="1" max="999" value="1" placeholder="Nhập số vòng lặp...">
                        <div class="tr-num-hint">Nhập số vòng lặp lại (Ví dụ: 2 sẽ đọc hết danh sách rồi tự động đọc lại lần 2). Mặc định là 1 vòng.</div>
                    </div>
                    <div class="tr-num-footer">
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-loop-cancel">Huỷ</button>
                        <button class="tr-btn tr-btn-primary tr-btn-sm" id="tr-loop-ok">Bắt đầu cào</button>
                    </div>`;
        overlay.appendChild(box);
        document.body.appendChild(overlay);
        const input = box.querySelector('#tr-loop-input');
        input.focus(); input.select();
        box.querySelector('#tr-loop-cancel').addEventListener('click', () => { overlay.remove(); resolve(null); });
        box.querySelector('#tr-loop-ok').addEventListener('click', () => {
            const val = parseInt(input.value, 10);
            if (!val || val < 1) { input.focus(); return; }
            overlay.remove(); resolve(val);
        });
        input.addEventListener('keydown', (e) => { if (e.key === 'Enter') box.querySelector('#tr-loop-ok').click(); });
        overlay.addEventListener('click', (e) => { if (e.target === overlay) { overlay.remove(); resolve(null); } });
    });

    const showBookStartGridDialog = (books) => new Promise((resolve) => {
        const old = document.getElementById('tr-grid-start-overlay');
        if (old) old.remove();

        const overlay = document.createElement('div');
        overlay.id = 'tr-grid-start-overlay';
        const box = document.createElement('div');
        box.className = 'tr-grid-start-box';
        box.innerHTML = `
                <div class="tr-grid-start-header">
                    <div class="tr-grid-start-icon">
                        <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="#0B57D0" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                            <rect x="3" y="3" width="7" height="7" rx="2"></rect>
                            <rect x="14" y="3" width="7" height="7" rx="2"></rect>
                            <rect x="14" y="14" width="7" height="7" rx="2"></rect>
                            <rect x="3" y="14" width="7" height="7" rx="2"></rect>
                        </svg>
                    </div>
                    <div>
                        <div class="tr-grid-start-title">Chọn cuốn sách bắt đầu chạy</div>
                        <div class="tr-grid-start-sub">Bấm chọn 1 ô vuông sách bất kỳ để bắt đầu chạy từ cuốn đó (${books.length} cuốn)</div>
                    </div>
                </div>
                <div class="tr-grid-start-search-wrap">
                    <input type="text" id="tr-grid-start-search" placeholder="🔍 Tìm kiếm tên sách nhanh...">
                </div>
                <div class="tr-grid-start-container" id="tr-grid-items">
                    ${books.map((b, idx) => `
                        <div class="tr-book-square" data-idx="${idx}" title="${b.title ? b.title.replace(/"/g, '&quot;') : ''}">
                            <div class="tr-book-square-num">#${idx + 1}</div>
                            <div class="tr-book-square-icon">
                                <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="#0B57D0" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>
                                    <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path>
                                </svg>
                            </div>
                            <div class="tr-book-square-title">${b.title || 'Cuốn sách #' + (idx + 1)}</div>
                        </div>
                    `).join('')}
                </div>
                <div class="tr-grid-start-footer">
                    <div class="tr-grid-start-preview" id="tr-grid-preview">Di chuột vào ô vuông để xem tên sách</div>
                    <div class="tr-grid-start-actions">
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-grid-back">Quay lại</button>
                        <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-grid-cancel">Huỷ</button>
                    </div>
                </div>
            `;
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        const previewEl = box.querySelector('#tr-grid-preview');
        const searchInput = box.querySelector('#tr-grid-start-search');
        const squares = box.querySelectorAll('.tr-book-square');

        squares.forEach(sq => {
            sq.addEventListener('mouseenter', () => {
                const idx = parseInt(sq.dataset.idx, 10);
                const book = books[idx];
                if (book) previewEl.textContent = '#' + (idx + 1) + ': ' + (book.title || 'Không có tên');
            });
            sq.addEventListener('click', () => {
                const idx = parseInt(sq.dataset.idx, 10);
                overlay.remove();
                resolve({
                    startIndex: idx,
                    books: books.slice(idx),
                    label: 'Bắt đầu từ cuốn #' + (idx + 1)
                });
            });
        });

        searchInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            squares.forEach(sq => {
                const idx = parseInt(sq.dataset.idx, 10);
                const b = books[idx];
                const title = (b.title || '').toLowerCase();
                const numStr = (idx + 1).toString();
                if (!query || title.includes(query) || numStr === query) {
                    sq.style.display = 'flex';
                } else {
                    sq.style.display = 'none';
                }
            });
        });

        box.querySelector('#tr-grid-back').addEventListener('click', () => {
            overlay.remove();
            resolve('back');
        });
        box.querySelector('#tr-grid-cancel').addEventListener('click', () => {
            overlay.remove();
            resolve(null);
        });
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                overlay.remove();
                resolve(null);
            }
        });
    });

    const showDirectionDialog = async (books) => {
        while (true) {
            const chosen = await new Promise((resolve) => {
                const old = document.getElementById('tr-direction-overlay');
                if (old) old.remove();

                const overlay = document.createElement('div');
                overlay.id = 'tr-direction-overlay';
                const box = document.createElement('div');
                box.className = 'tr-direction-box';
                box.innerHTML = `
                        <div class="tr-direction-header">
                            <div class="tr-direction-icon">
                                <svg viewBox="0 0 24 24" width="30" height="30" fill="#0B57D0">
                                    <path d="M16 17.01V10h-2v7.01h-3L15 21l4-3.99h-3zM9 3L5 6.99h3V14h2V6.99h3L9 3z"/>
                                </svg>
                            </div>
                            <div class="tr-direction-title">Chọn hướng bắt đầu đọc</div>
                            <div class="tr-direction-sub">Danh sách gồm <strong>${books.length} cuốn sách</strong>. Cậu muốn bắt đầu chạy theo hướng nào?</div>
                        </div>
                        <div class="tr-direction-options">
                            <div class="tr-direction-card" id="tr-dir-bottom-up">
                                <div class="tr-direction-card-icon tr-dir-primary">
                                    <svg viewBox="0 0 24 24" width="22" height="22" fill="currentColor">
                                        <path d="M4 12l1.41 1.41L11 7.83V20h2V7.83l5.58 5.59L20 12l-8-8-8 8z"/>
                                    </svg>
                                </div>
                                <div class="tr-direction-card-body">
                                    <div class="tr-direction-card-title-row">
                                        <span class="tr-direction-card-title">Từ dưới lên trên</span>
                                        <span class="tr-direction-badge">Khuyên dùng</span>
                                    </div>
                                    <div class="tr-direction-card-desc">Đọc từ cuốn cuối danh sách (cũ nhất) ngược dần lên cuốn đầu tiên.</div>
                                </div>
                                <div class="tr-direction-card-arrow">
                                    <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                                        <path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/>
                                    </svg>
                                </div>
                            </div>
                            <div class="tr-direction-card" id="tr-dir-top-down">
                                <div class="tr-direction-card-icon">
                                    <svg viewBox="0 0 24 24" width="22" height="22" fill="currentColor">
                                        <path d="M20 12l-1.41-1.41L13 16.17V4h-2v12.17l-5.58-5.59L4 12l8 8 8-8z"/>
                                    </svg>
                                </div>
                                <div class="tr-direction-card-body">
                                    <div class="tr-direction-card-title-row">
                                        <span class="tr-direction-card-title">Từ trên xuống dưới</span>
                                    </div>
                                    <div class="tr-direction-card-desc">Đọc từ cuốn đầu tiên (mới nhất) xuôi dần xuống các cuốn bên dưới.</div>
                                </div>
                                <div class="tr-direction-card-arrow">
                                    <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                                        <path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/>
                                    </svg>
                                </div>
                            </div>
                        </div>
                        <div class="tr-direction-footer">
                            <button class="tr-btn tr-btn-secondary tr-btn-sm" id="tr-dir-cancel">Huỷ</button>
                        </div>
                    `;
                overlay.appendChild(box);
                document.body.appendChild(overlay);

                box.querySelector('#tr-dir-bottom-up').addEventListener('click', () => {
                    overlay.remove();
                    resolve({
                        mode: 'bottom_up',
                        books: books.slice().reverse(),
                        label: 'Từ dưới lên'
                    });
                });
                box.querySelector('#tr-dir-top-down').addEventListener('click', () => {
                    overlay.remove();
                    resolve({
                        mode: 'top_down',
                        books: books.slice(),
                        label: 'Từ trên xuống'
                    });
                });
                box.querySelector('#tr-dir-cancel').addEventListener('click', () => {
                    overlay.remove();
                    resolve(null);
                });
                overlay.addEventListener('click', (e) => {
                    if (e.target === overlay) {
                        overlay.remove();
                        resolve(null);
                    }
                });
            });

            if (chosen === null) return null;
            return chosen;
        }
    };

    const injectStyles = () => {
        const style = document.createElement('style');
        const bgUrl = getBgUrl();
        style.textContent = `
        @keyframes tr-cd-fade-in { from { opacity:0 } to { opacity:1 } }
        @keyframes tr-cd-pop { 0% { transform: scale(.85); opacity:0 } 100% { transform: scale(1); opacity:1 } }
        @keyframes tr-cd-num-pop { 0% { transform: scale(1) } 50% { transform: scale(1.1) } 100% { transform: scale(1) } }
        @keyframes tr-pulse-dot { 0%, 100% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.15); opacity: 0.7; } }
        @keyframes tr-step-glow { 0%, 100% { box-shadow: 0 0 0 0 rgba(11,87,208,0.5); } 50% { box-shadow: 0 0 0 6px rgba(11,87,208,0); } }
        @keyframes tr-modal-in { from { opacity:0; transform:scale(.9) } to { opacity:1; transform:scale(1) } }
        @keyframes tr-panel-push-up {
            0% { transform: translateY(30px) scale(0.96); opacity: 0.85; }
            100% { transform: translateY(0) scale(1); opacity: 1; }
        }

        canvas, .df-container, .df-book-stage, .df-ui-wrapper, div[class*="df-"] {
            transform: translate3d(0, 0, 0) !important;
            -webkit-transform: translate3d(0, 0, 0) !important;
            will-change: transform, opacity !important;
            backface-visibility: hidden !important;
            perspective: 1000px !important;
        }

        .tr-panel {
            position: fixed; bottom: 20px; left: 20px;
            width: 440px;
            max-height: 94vh;
            background-image: url('${bgUrl}');
            background-size: cover;
            background-position: 70% center;
            background-repeat: no-repeat;
            border: 1px solid rgba(255,255,255,0.6);
            border-radius: 28px;
            box-shadow:
                0 16px 44px rgba(0,0,0,0.24),
                0 2px 6px rgba(0,0,0,0.08),
                inset 0 1px 0 rgba(255,255,255,0.6);
            z-index: 2147483646;
            font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif;
            color: #1A1C1E; overflow: hidden;
            transition: width 0.45s cubic-bezier(0.34, 1.2, 0.64, 1),
                        border-radius 0.4s cubic-bezier(0.2, 0.9, 0.3, 1),
                        box-shadow 0.35s ease,
                        transform 0.35s cubic-bezier(0.34, 1.2, 0.64, 1);
            display: flex;
            flex-direction: column;
            transform-origin: bottom left;
            will-change: width, border-radius, transform;
        }
        .tr-panel:hover {
            box-shadow:
                0 20px 48px rgba(0,0,0,0.26),
                0 4px 10px rgba(0,0,0,0.08),
                inset 0 1px 0 rgba(255,255,255,0.7);
        }
        .tr-panel.tr-panel-expanded {
            min-height: auto;
            animation: tr-panel-push-up 0.5s cubic-bezier(0.2, 0.8, 0.2, 1);
            transform-origin: bottom center;
        }
        .tr-panel.tr-collapsed {
            width: 250px;
            min-height: auto !important;
            border-radius: 100px;
            box-shadow: 0 12px 32px rgba(0,0,0,0.2), inset 0 1px 0 rgba(255,255,255,0.7);
            transition: width 0.4s cubic-bezier(0.2, 0.9, 0.3, 1),
                        border-radius 0.4s cubic-bezier(0.2, 0.9, 0.3, 1),
                        box-shadow 0.35s ease,
                        transform 0.3s cubic-bezier(0.2, 0.9, 0.3, 1);
        }
        .tr-panel.tr-collapsed:hover {
            transform: translateY(-3px) scale(1.025);
            box-shadow: 0 18px 42px rgba(0,0,0,0.28), inset 0 1px 0 rgba(255,255,255,0.85);
        }

        .tr-header {
            display: flex; align-items: center; gap: 12px;
            padding: 18px 22px;
            background: transparent;
            cursor: pointer; user-select: none;
            position: relative;
            border-bottom: 1px solid rgba(255,255,255,0.5);
            flex-shrink: 0;
            transition: padding 0.35s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.3s ease;
        }
        .tr-panel.tr-collapsed .tr-header {
            padding: 12px 18px;
            border-bottom-color: transparent;
        }
        .tr-header::after { display: none; }
        .tr-logo { width: 40px; height: 40px; object-fit: contain; border-radius: 50%; flex-shrink: 0; }
        .tr-title {
            flex: 1; font-size: 16.5px; font-weight: 800; letter-spacing: -0.2px; line-height: 1.25;
            color: #ffffff !important;
            text-shadow: 0 1px 3px rgba(0, 0, 0, 0.8);
            transition: font-size 0.3s ease;
        }
        .tr-panel.tr-collapsed .tr-title {
            font-size: 15px;
        }
        .tr-subtitle {
            font-size: 11.5px; font-weight: 700; display: block; margin-top: 2px;
            color: #0284C7; font-style: normal;
            text-shadow: 0 1px 2px rgba(255,255,255,0.9);
            transition: font-size 0.3s ease;
        }
        .tr-panel.tr-collapsed .tr-subtitle {
            font-size: 11px;
        }
        .tr-collapse {
            width: 34px; height: 34px; border-radius: 50%;
            background: rgba(255,255,255,0.7); border: 1px solid rgba(255,255,255,0.9);
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            cursor: pointer; display: flex; align-items: center; justify-content: center;
            color: #0F172A;
            transition: background .2s ease, transform .4s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow .2s ease;
            flex-shrink: 0;
        }
        .tr-collapse:hover {
            background: rgba(255,255,255,0.95);
            transform: scale(1.12);
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        }
        .tr-collapse svg {
            width: 18px; height: 18px;
            transition: transform 0.45s cubic-bezier(0.34, 1.56, 0.64, 1);
        }
        .tr-panel.tr-collapsed .tr-collapse svg {
            transform: rotate(180deg) scale(0.92);
        }

        .tr-body {
            padding: 18px 20px 20px;
            display: flex; flex-direction: column; gap: 14px;
            flex: 1; min-height: 0; overflow-y: auto;
            background: transparent;
            max-height: 600px;
            opacity: 1;
            transform: translateY(0);
            transition: max-height 0.45s cubic-bezier(0.34, 1.15, 0.64, 1),
                        opacity 0.3s cubic-bezier(0.2, 0.8, 0.2, 1),
                        transform 0.4s cubic-bezier(0.34, 1.15, 0.64, 1),
                        padding 0.35s cubic-bezier(0.16, 1, 0.3, 1);
            will-change: max-height, opacity, transform;
        }
        .tr-body::-webkit-scrollbar { width: 6px }
        .tr-body::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.5); border-radius: 3px }
        .tr-body::-webkit-scrollbar-track { background: transparent }
        .tr-footer-note {
            color: rgba(255, 255, 255, 0.9);
            font-size: 7.5px;
            font-weight: 700;
            line-height: 1.35;
            letter-spacing: 0.2px;
            text-align: right;
            text-shadow: 0 1px 2px rgba(0, 0, 0, 0.7);
            margin-top: -6px;
            padding-top: 2px;
        }
        .tr-view-count-debug {
            display: none !important;
            flex-shrink: 0;
            padding: 10px 12px;
            border: 1px solid rgba(147, 197, 253, 0.42);
            border-radius: 12px;
            background: rgba(15, 23, 42, 0.82);
            color: #F8FAFC;
            font-size: 11px;
            line-height: 1.45;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.18);
        }
        .tr-view-count-debug-head {
            color: #BFDBFE;
            font-size: 10px;
            font-weight: 800;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            margin-bottom: 4px;
        }
        .tr-view-count-debug-book {
            overflow: hidden;
            margin-bottom: 6px;
            color: #FFFFFF;
            font-weight: 650;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .tr-view-count-debug-row {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            color: rgba(226, 232, 240, 0.78);
        }
        .tr-view-count-debug-row strong {
            color: #FFFFFF;
            font-family: ui-monospace, SFMono-Regular, Consolas, monospace;
            font-weight: 700;
            text-align: right;
        }
        .tr-view-count-debug-status {
            margin-top: 6px;
            color: #BFDBFE;
            font-size: 10px;
            overflow-wrap: anywhere;
        }
        .tr-panel.tr-collapsed .tr-body {
            max-height: 0;
            opacity: 0;
            transform: translateY(12px) scale(0.98);
            padding-top: 0;
            padding-bottom: 0;
            overflow: hidden;
            pointer-events: none;
            transition: max-height 0.35s cubic-bezier(0.2, 0.9, 0.3, 1),
                        opacity 0.2s ease,
                        transform 0.3s ease,
                        padding 0.35s ease;
        }

        .tr-slider-stepper {
            width: 100%;
            padding: 6px 0 2px;
            margin-top: auto;
            flex-shrink: 0;
        }
        .tr-stepper-meta {
            display: flex;
            justify-content: flex-end;
            min-height: 14px;
            padding-top: 4px;
        }
        .tr-stepper-tag {
            color: #0B57D0;
            font-size: 10px;
            font-weight: 700;
            font-variant-numeric: tabular-nums;
            text-shadow: 0 1px 2px rgba(255, 255, 255, 0.8);
        }
        .tr-stepper-track-wrap {
            position: relative;
            width: 100%;
            height: 8px;
            background: rgba(203, 213, 225, 0.55);
            border-radius: 100px;
            overflow: hidden;
            box-sizing: border-box;
            box-shadow: inset 0 1px 2px rgba(0,0,0,0.06);
        }
        .tr-stepper-active-line {
            height: 100%;
            background: #0B57D0;
            border-radius: 100px;
            width: 0%;
            min-width: 0%;
            transition: width .35s cubic-bezier(.2,0,0,1);
            box-shadow: none;
        }

        #tr-scrape-overlay {
            position: fixed; inset: 0;
            background: rgba(0, 0, 0, 0.45);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            z-index: 2147483647;
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            gap: 0;
            font-family: 'Roboto', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            animation: tr-cd-fade-in .25s ease-out;
            -webkit-font-smoothing: antialiased;
            pointer-events: none;
        }
        .tr-scrape-message {
            color: #FFFFFF !important;
            font-size: 18px; font-weight: 500; letter-spacing: .2px;
            text-align: center; padding: 0 20px;
            text-shadow: 0 1px 6px rgba(0, 0, 0, 0.7);
        }

        .tr-perf-card {
            display: flex; align-items: center; justify-content: space-between;
            padding: 10px 14px;
            background: rgba(255, 255, 255, 0.65);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.8);
            border-radius: 18px;
            cursor: pointer; user-select: none;
            transition: all .25s cubic-bezier(.2,0,0,1);
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.03);
            flex-shrink: 0;
        }
        .tr-perf-card:hover {
            background: rgba(255, 255, 255, 0.9);
            border-color: #0B57D0;
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(11, 87, 208, 0.12);
        }
        .tr-perf-card.active {
            background: linear-gradient(135deg, rgba(255, 247, 237, 0.95), rgba(254, 243, 199, 0.92));
            border-color: rgba(245, 158, 11, 0.7);
            box-shadow: 0 6px 20px rgba(245, 158, 11, 0.18);
        }
        .tr-perf-info {
            display: flex; align-items: center; gap: 10px;
        }
        .tr-perf-icon-wrap {
            width: 34px; height: 34px; border-radius: 50%;
            background: rgba(11, 87, 208, 0.1); color: #0B57D0;
            display: flex; align-items: center; justify-content: center;
            transition: all .25s ease; flex-shrink: 0;
        }
        .tr-perf-card.active .tr-perf-icon-wrap {
            background: linear-gradient(135deg, #FF6B00, #F59E0B);
            color: #FFFFFF;
            box-shadow: 0 2px 10px rgba(245, 158, 11, 0.45);
        }
        .tr-perf-text {
            display: flex; flex-direction: column; gap: 2px;
        }
        .tr-perf-title {
            font-size: 13px; font-weight: 700; color: #1E293B; letter-spacing: -0.1px;
        }
        .tr-perf-card.active .tr-perf-title {
            color: #B45309;
        }
        .tr-perf-desc {
            font-size: 11px; color: #64748B; font-weight: 500;
        }
        .tr-perf-card.active .tr-perf-desc {
            color: #D97706;
        }
        .tr-perf-switch {
            width: 42px; height: 24px; border-radius: 100px;
            background: #CBD5E1; position: relative;
            transition: background .25s cubic-bezier(.2,0,0,1);
            flex-shrink: 0;
        }
        .tr-perf-thumb {
            width: 18px; height: 18px; border-radius: 50%;
            background: #FFFFFF; position: absolute;
            top: 3px; left: 3px;
            transition: transform .25s cubic-bezier(.2,0,0,1);
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.25);
        }
        .tr-perf-card.active .tr-perf-switch {
            background: #F59E0B;
        }
        .tr-perf-card.active .tr-perf-thumb {
            transform: translateX(18px);
        }

        .tr-toast {
            position: fixed; top: 24px; left: 50%;
            transform: translate(-50%, -20px) scale(0.95);
            background: rgba(15, 23, 42, 0.94);
            color: #FFFFFF;
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 100px;
            padding: 12px 24px;
            display: flex; align-items: center; gap: 10px;
            font-size: 13.5px; font-weight: 600;
            font-family: 'Roboto', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            box-shadow: 0 12px 36px rgba(0, 0, 0, 0.4), 0 0 20px rgba(11, 87, 208, 0.3);
            backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
            z-index: 2147483647; opacity: 0; pointer-events: none;
            transition: all .3s cubic-bezier(.2,0,0,1);
            -webkit-font-smoothing: antialiased;
        }
        .tr-toast.tr-toast-show {
            opacity: 1; transform: translate(-50%, 0) scale(1); pointer-events: auto;
        }

        .tr-eta { display: none !important; }
        .tr-eta.tr-on { display: none !important; }

        #tr-page-blocker {
            display: none;
            position: fixed;
            inset: 0;
            z-index: 2147483640;
            background: transparent;
            cursor: not-allowed;
            user-select: none;
            -webkit-user-select: none;
        }

        .tr-btn-group { margin-top: auto; display: flex; flex-direction: column; gap: 10px; flex-shrink: 0; }
        .tr-btn-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            width: 100%;
        }
        .tr-btn-row .tr-btn {
            padding: 10px 10px;
            font-size: 12px;
            justify-content: center;
            white-space: nowrap;
            gap: 6px;
        }
        .tr-btn-row .tr-btn svg { width: 15px; height: 15px; }

        .tr-btn {
            border: 1px solid rgba(255,255,255,0.55);
            border-radius: 100px;
            padding: 13px 20px;
            font-size: 13.5px; font-weight: 600;
            font-family: inherit; cursor: pointer;
            display: flex; align-items: center; justify-content: center; gap: 10px;
            transition: background .2s, box-shadow .2s, opacity .2s, transform .1s, border-color .2s;
            position: relative; overflow: hidden; flex-shrink: 0;
        }
        .tr-btn-sm { padding: 10px 16px; font-size: 12.5px }
        .tr-btn svg { width: 17px; height: 17px; flex-shrink: 0; position: relative; z-index: 1 }
        .tr-btn span { position: relative; z-index: 1 }
        .tr-btn:disabled { opacity: .45; cursor: not-allowed }
        .tr-btn:active:not(:disabled) { transform: scale(.98) }
        .tr-btn-primary { background: rgba(11,87,208,0.85); color: #fff; border-color: rgba(255,255,255,0.45); box-shadow: 0 6px 20px rgba(11,87,208,0.38), inset 0 1px 0 rgba(255,255,255,0.4); }
        .tr-btn-primary:hover:not(:disabled) { background: rgba(10,76,184,0.95); }
        .tr-btn-primary svg { filter: drop-shadow(0 0 4px rgba(255,255,255,0.6)); }
        .tr-btn-secondary { background: rgba(255,255,255,0.4); color: #0B57D0; border-color: rgba(255,255,255,0.6); box-shadow: 0 3px 12px rgba(0,0,0,0.08), inset 0 1px 0 rgba(255,255,255,0.7); }
        .tr-btn-secondary:hover:not(:disabled) { background: rgba(255,255,255,0.6); border-color: rgba(255,255,255,0.85); }
        .tr-btn-danger { background: rgba(186,26,26,0.85); color: #fff; border-color: rgba(255,255,255,0.45); box-shadow: 0 6px 20px rgba(186,26,26,0.38), inset 0 1px 0 rgba(255,255,255,0.4); }
        .tr-btn-danger:hover:not(:disabled) { background: rgba(160,24,24,0.95); }
        .tr-btn-outline-danger { background: rgba(255,255,255,0.35); color: #BA1A1A; border-color: rgba(186,26,26,0.25); box-shadow: 0 2px 8px rgba(186,26,26,0.06), inset 0 1px 0 rgba(255,255,255,0.5); }
        .tr-btn-outline-danger:hover:not(:disabled) { background: rgba(186,26,26,0.08); border-color: rgba(186,26,26,0.4); }

        #tr-countdown-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 2147483647; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 14px; animation: tr-cd-fade-in .3s ease-out; transition: opacity .25s ease-out; font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif; pointer-events: none; }
        #tr-countdown-overlay.tr-cd-out { opacity: 0; pointer-events: none }
        .tr-cd-label { color: rgba(255,255,255,0.9); font-size: 13px; font-weight: 400; letter-spacing: .3px; animation: tr-cd-pop .4s ease-out; text-align: center; padding: 0 20px; text-shadow: 0 1px 4px rgba(0,0,0,0.6) }
        .tr-cd-num { color: #fff; font-size: 56px; font-weight: 300; line-height: 1; animation: tr-cd-num-pop .5s ease-out; font-variant-numeric: tabular-nums; text-shadow: 0 2px 12px rgba(0,0,0,0.7) }

        #tr-report-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 2147483647; display: flex; align-items: center; justify-content: center }
        .tr-report-close { position: absolute; top: 20px; right: 20px; width: 42px; height: 42px; border-radius: 50%; background: #BA1A1A; color: #fff; border: none; font-size: 20px; font-weight: 700; cursor: pointer; z-index: 10 }
        .tr-report-iframe { width: 92vw; height: 92vh; border: none; border-radius: 14px; background: #0F172A }

        #tr-picker-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.6); z-index: 2147483647; display: flex; align-items: center; justify-content: center; font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif; }
        .tr-picker-box { background: rgba(254,247,255,0.97); border: 1px solid rgba(255,255,255,0.6); border-radius: 28px; width: 90vw; max-width: 720px; max-height: 88vh; display: flex; flex-direction: column; box-shadow: 0 24px 60px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.7); animation: tr-modal-in .35s cubic-bezier(.2,0,0,1); overflow: hidden; }
        .tr-picker-header { padding: 20px 24px 12px; border-bottom: 1px solid rgba(226,232,240,0.8) }
        .tr-picker-title { font-size: 17px; font-weight: 700; color: #1A1C1E }
        .tr-picker-sub { font-size: 12px; color: #64748B; margin-top: 4px }
        .tr-picker-controls { padding: 16px 24px; background: rgba(248,250,252,0.7); border-bottom: 1px solid rgba(226,232,240,0.8); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; }
        .tr-picker-ctrl { display: flex; align-items: center; gap: 10px; font-size: 13px; color: #334155 }
        .tr-picker-ctrl label { font-weight: 600 }
        .tr-picker-ctrl input[type="number"] { width: 90px; padding: 8px 12px; border-radius: 10px; border: 1px solid #CBD5E1; font-family: inherit; font-size: 13px; outline: none; background: rgba(255,255,255,0.9); }
        .tr-picker-ctrl input[type="number"]:focus { border-color: #0B57D0; box-shadow: 0 0 0 3px rgba(11,87,208,0.15) }
        .tr-picker-hint { font-size: 11.5px; color: #94A3B8 }
        .tr-picker-actions { display: flex; gap: 8px }
        .tr-picker-mini { padding: 7px 14px; border-radius: 100px; background: rgba(226,232,240,0.85); color: #334155; border: 1px solid rgba(255,255,255,0.6); cursor: pointer; font-family: inherit; font-size: 12px; font-weight: 600; transition: background .15s; }
        .tr-picker-mini:hover { background: rgba(203,213,225,0.95) }
        .tr-picker-list { flex: 1; overflow-y: auto; padding: 8px 12px }
        .tr-picker-item { display: flex; align-items: flex-start; gap: 12px; padding: 10px 14px; border-radius: 12px; cursor: pointer; transition: background .12s; }
        .tr-picker-item:hover { background: rgba(241,245,249,0.8) }
        .tr-picker-item input[type="checkbox"] { width: 18px; height: 18px; margin-top: 2px; cursor: pointer; accent-color: #0B57D0; flex-shrink: 0; }
        .tr-picker-item-info { flex: 1; min-width: 0 }
        .tr-picker-item-title { font-size: 13px; font-weight: 600; color: #1A1C1E; line-height: 1.35; word-break: break-word }
        .tr-picker-item-meta { font-size: 11px; color: #64748B; margin-top: 3px; font-weight: 500 }
        .tr-picker-item.tr-read .tr-picker-item-title { color: #64748B }
        .tr-picker-item.tr-read .tr-picker-item-meta { color: #16A34A }
        .tr-picker-footer { padding: 14px 24px; border-top: 1px solid rgba(226,232,240,0.8); display: flex; justify-content: space-between; align-items: center; background: rgba(248,250,252,0.7); }
        .tr-picker-count { font-size: 13px; font-weight: 700; color: #0B57D0 }
        .tr-picker-footer > div:last-child { display: flex; gap: 10px }

        #tr-mode-overlay, #tr-num-overlay, #tr-loop-overlay {
            position: fixed; inset: 0; background: rgba(15, 23, 42, 0.45);
            backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
            z-index: 2147483647; display: flex; align-items: center; justify-content: center;
            font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif;
            animation: tr-cd-fade-in .2s ease-out;
        }
        .tr-mode-box {
            background: #FFFFFF;
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 28px; width: 92vw; max-width: 480px;
            box-shadow: 0 24px 60px rgba(15, 23, 42, 0.22), 0 4px 16px rgba(15, 23, 42, 0.06);
            animation: tr-modal-in .3s cubic-bezier(.2,0,0,1);
            overflow: hidden; display: flex; flex-direction: column;
        }
        .tr-num-box {
            background: #FFFFFF;
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 28px; width: 90vw; max-width: 440px;
            box-shadow: 0 24px 60px rgba(15, 23, 42, 0.22), 0 4px 16px rgba(15, 23, 42, 0.06);
            animation: tr-modal-in .3s cubic-bezier(.2,0,0,1);
            overflow: hidden;
        }
        .tr-mode-header { padding: 24px 24px 14px; text-align: center; }
        .tr-num-header { padding: 24px 24px 16px; text-align: center; }
        .tr-mode-icon {
            width: 52px; height: 52px; border-radius: 18px;
            background: linear-gradient(135deg, #0B57D0, #0842A0);
            color: #FFFFFF; display: flex; align-items: center; justify-content: center;
            margin: 0 auto 12px; box-shadow: 0 8px 20px rgba(11, 87, 208, 0.3);
            animation: tr-lightning-box-glow 1.8s ease-in-out infinite;
        }
        .tr-mode-icon svg {
            animation: tr-lightning-flash 1.8s ease-in-out infinite;
        }
        @keyframes tr-lightning-box-glow {
            0%, 100% { box-shadow: 0 8px 20px rgba(11, 87, 208, 0.3); }
            30% { box-shadow: 0 8px 28px rgba(11, 87, 208, 0.65), 0 0 12px rgba(96, 165, 250, 0.5); }
            60% { box-shadow: 0 8px 22px rgba(11, 87, 208, 0.4); }
        }
        @keyframes tr-lightning-flash {
            0%, 100% { opacity: 1; transform: scale(1); filter: drop-shadow(0 0 6px rgba(255,255,255,0.9)); }
            12% { opacity: 0.2; transform: scale(0.85); filter: none; }
            24% { opacity: 1; transform: scale(1.15) rotate(-3deg); filter: drop-shadow(0 0 12px rgba(255,255,255,1)); }
            36% { opacity: 0.3; transform: scale(0.9); filter: none; }
            48% { opacity: 1; transform: scale(1.08) rotate(2deg); filter: drop-shadow(0 0 10px rgba(255,255,255,0.9)); }
            65% { opacity: 0.85; transform: scale(1); filter: drop-shadow(0 0 5px rgba(255,255,255,0.7)); }
        }
        .tr-num-icon { font-size: 36px; margin-bottom: 8px; }
        .tr-mode-title { font-size: 19px; font-weight: 700; color: #1A1C1E; letter-spacing: -0.2px; }
        .tr-num-title { font-size: 18px; font-weight: 700; color: #1A1C1E; }
        .tr-mode-sub { font-size: 12.5px; color: #64748B; margin-top: 4px; }
        .tr-num-sub { font-size: 12.5px; color: #64748B; margin-top: 4px; }
        .tr-mode-options { padding: 4px 16px 12px; display: flex; flex-direction: column; gap: 9px; }
        .tr-mode-card {
            display: flex; align-items: center; gap: 14px; padding: 13px 16px;
            border-radius: 18px; background: #F8FAFD;
            border: 1.5px solid #E2E8F0; cursor: pointer;
            transition: all .2s cubic-bezier(0.2, 0, 0, 1);
        }
        .tr-mode-card:hover {
            background: #FFFFFF; border-color: #0B57D0;
            transform: translateY(-2px);
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
        }
        .tr-mode-card:active { transform: scale(0.985); }
        .tr-mode-card-icon {
            flex-shrink: 0; width: 44px; height: 44px; border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            transition: transform .2s cubic-bezier(0.2, 0, 0, 1);
        }
        .tr-mode-card:hover .tr-mode-card-icon { transform: scale(1.08); }
        .tr-mode-blue { background: #EEF2FF; color: #3B82F6; border: 1px solid #C7D2FE; }
        .tr-mode-green { background: #ECFDF5; color: #10B981; border: 1px solid #A7F3D0; }
        .tr-mode-purple { background: #FAF5FF; color: #8B5CF6; border: 1px solid #DDD6FE; }
        .tr-mode-cyan { background: #F0F9FF; color: #0284C7; border: 1px solid #BAE6FD; }
        .tr-mode-card-body { flex: 1; min-width: 0; }
        .tr-mode-card-title-row { display: flex; align-items: center; gap: 8px; margin-bottom: 2px; }
        .tr-mode-card-title { font-size: 14px; font-weight: 700; color: #1E293B; }
        .tr-mode-pill { font-size: 10.5px; font-weight: 700; padding: 2px 7px; border-radius: 99px; letter-spacing: 0.2px; }
        .tr-pill-blue { background: #EEF2FF; color: #2563EB; }
        .tr-pill-green { background: #ECFDF5; color: #059669; }
        .tr-pill-purple { background: #FAF5FF; color: #7C3AED; }
        .tr-pill-cyan { background: #F0F9FF; color: #0284C7; }
        .tr-mode-card-desc { font-size: 11.5px; color: #64748B; line-height: 1.4; }
        .tr-mode-card-arrow {
            color: #94A3B8; flex-shrink: 0; width: 28px; height: 28px;
            display: flex; align-items: center; justify-content: center;
            border-radius: 50%; background: transparent;
            transition: all .2s;
        }
        .tr-mode-card:hover .tr-mode-card-arrow {
            transform: translateX(3px); color: #0B57D0; background: #EEF2FF;
        }
        .tr-mode-footer {
            padding: 12px 20px 16px; border-top: 1px solid #F1F5F9;
            display: flex; justify-content: flex-end; background: #FAFAFA;
        }
        .tr-mode-info { font-size: 12px; color: #64748B; font-weight: 500; }
        .tr-num-input-wrap { padding: 8px 24px 16px; }
        .tr-num-input-wrap input { width: 100%; padding: 14px 18px; border-radius: 14px; border: 1.5px solid #CBD5E1; font-family: inherit; font-size: 16px; font-weight: 600; text-align: center; outline: none; background: rgba(255,255,255,0.9); transition: border-color .15s, box-shadow .15s; box-sizing: border-box; }
        .tr-num-input-wrap input:focus { border-color: #0B57D0; box-shadow: 0 0 0 3px rgba(11,87,208,0.15); }
        .tr-num-hint { font-size: 11.5px; color: #94A3B8; text-align: center; margin-top: 8px; }
        .tr-num-footer { padding: 14px 24px; border-top: 1px solid rgba(226,232,240,0.8); display: flex; justify-content: flex-end; gap: 10px; background: rgba(248,250,252,0.7); }

        #tr-blue-alert-overlay {
            position: fixed; inset: 0; background: rgba(15, 23, 42, 0.45);
            backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
            z-index: 2147483647; display: flex; align-items: center; justify-content: center;
            font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif;
            animation: tr-cd-fade-in .2s ease-out;
        }
        .tr-blue-alert-box {
            background: #FFFFFF;
            border: 1px solid rgba(199, 210, 254, 0.9);
            border-radius: 28px; width: 90vw; max-width: 420px;
            box-shadow: 0 24px 60px rgba(11, 87, 208, 0.22), 0 4px 16px rgba(15, 23, 42, 0.06);
            padding: 24px 24px 20px; text-align: center;
            animation: tr-modal-in .3s cubic-bezier(.2,0,0,1);
        }
        .tr-blue-alert-icon {
            width: 54px; height: 54px; border-radius: 18px;
            background: #EEF2FF; color: #0B57D0; border: 1px solid #C7D2FE;
            display: flex; align-items: center; justify-content: center;
            margin: 0 auto 12px; box-shadow: 0 8px 20px rgba(11, 87, 208, 0.2);
        }
        .tr-blue-alert-title {
            font-size: 18px; font-weight: 700; color: #0F172A; margin-bottom: 6px;
        }
        .tr-blue-alert-body {
            font-size: 13px; color: #475569; line-height: 1.5; margin-bottom: 20px;
        }
        .tr-blue-alert-footer {
            display: flex; justify-content: center;
        }
        .tr-blue-alert-btn {
            min-width: 130px; border-radius: 14px; padding: 10px 24px; font-size: 14px; font-weight: 600;
        }

        #tr-confirm-overlay, #tr-summary-overlay {
            position: fixed; inset: 0; background: rgba(0,0,0,0.55);
            backdrop-filter: blur(5px); z-index: 2147483647;
            display: flex; align-items: center; justify-content: center;
            font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif;
            animation: tr-cd-fade-in .2s ease-out;
        }
        .tr-confirm-box {
            background: rgba(254,247,255,0.98);
            border: 1px solid rgba(255,255,255,0.7);
            border-radius: 24px; width: 90vw; max-width: 420px;
            box-shadow: 0 24px 60px rgba(0,0,0,0.35);
            padding: 24px; animation: tr-modal-in .25s cubic-bezier(.2,0,0,1);
        }
        .tr-confirm-header { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }
        .tr-confirm-icon { font-size: 28px; }
        .tr-confirm-title { font-size: 17px; font-weight: 700; color: #1A1C1E; }
        .tr-confirm-body { font-size: 13.5px; color: #475569; line-height: 1.5; margin-bottom: 20px; }
        .tr-confirm-footer { display: flex; justify-content: flex-end; gap: 10px; }

        .tr-summary-box {
            background: rgba(254,247,255,0.98);
            border: 1px solid rgba(255,255,255,0.7);
            border-radius: 28px; width: 90vw; max-width: 520px; max-height: 85vh;
            display: flex; flex-direction: column;
            box-shadow: 0 28px 70px rgba(0,0,0,0.4);
            animation: tr-modal-in .3s cubic-bezier(.2,0,0,1); overflow: hidden;
        }
        .tr-summary-header { padding: 22px 24px 14px; text-align: center; }
        .tr-summary-icon { font-size: 38px; margin-bottom: 6px; }
        .tr-summary-title { font-size: 19px; font-weight: 700; color: #1A1C1E; }
        .tr-summary-sub { font-size: 12.5px; color: #64748B; margin-top: 4px; }
        .tr-summary-stats {
            display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px;
            padding: 0 22px 20px;
        }
        .tr-summary-stat-card {
            padding: 14px 10px; border-radius: 18px; text-align: center;
            background: rgba(255,255,255,0.85); border: 1px solid rgba(226,232,240,0.8);
        }
        .tr-summary-stat-val { font-size: 22px; font-weight: 800; line-height: 1.2; }
        .tr-summary-stat-lbl { font-size: 11.5px; font-weight: 600; color: #64748B; margin-top: 3px; }
        .stat-success .tr-summary-stat-val { color: #16A34A; }
        .stat-time .tr-summary-stat-val { color: #0B57D0; font-size: 19px; }

        .tr-summary-list-title {
            padding: 6px 22px; font-size: 12px; font-weight: 700; color: #475569;
            text-transform: uppercase; letter-spacing: 0.5px;
        }
        .tr-summary-list {
            flex: 1; overflow-y: auto; padding: 4px 20px 14px;
            display: flex; flex-direction: column; gap: 8px;
            max-height: 240px;
        }
        .tr-summary-item {
            display: flex; align-items: flex-start; gap: 10px;
            padding: 9px 12px; border-radius: 12px;
            background: rgba(255,255,255,0.85); border: 1px solid rgba(226,232,240,0.8);
        }
        .tr-summary-item-idx {
            font-size: 11px; font-weight: 700; color: #94A3B8;
            min-width: 18px; padding-top: 2px;
        }
        .tr-summary-item-content { flex: 1; min-width: 0; }
        .tr-summary-item-title {
            font-size: 12.5px; font-weight: 600; color: #1E293B;
            line-height: 1.35; word-break: break-word;
        }
        .tr-summary-item-reason { margin-top: 4px; }
        .tr-summary-badge {
            display: inline-block; font-size: 10.5px; font-weight: 600;
            padding: 2px 8px; border-radius: 100px;
        }
        .badge-teacher { background: #F5F3FF; color: #7C3AED; border: 1px solid #DDD6FE; }
        .badge-read { background: #ECFDF5; color: #059669; border: 1px solid #A7F3D0; }
        .badge-warn { background: #FFFBEB; color: #D97706; border: 1px solid #FDE68A; }
        .badge-general { background: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; }

        .tr-summary-empty {
            padding: 20px; text-align: center; color: #059669; font-size: 13px; font-weight: 500;
            background: #ECFDF5; margin: 8px 20px; border-radius: 14px; border: 1px solid #A7F3D0;
        }
        .tr-summary-footer {
            padding: 14px 20px; border-top: 1px solid rgba(226,232,240,0.8);
            display: flex; justify-content: space-between; align-items: center;
            background: rgba(248,250,252,0.8);
        }

        /* ===== MATERIAL 3 VIEW COUNT SCRAPE POPUP OVERLAY ===== */
        #tr-view-scrape-overlay {
            position: fixed;
            inset: 0;
            background: rgba(15, 23, 42, 0.45);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            z-index: 2147483645;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: 'Roboto', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            animation: tr-cd-fade-in 0.25s ease-out;
            padding: 20px;
            box-sizing: border-box;
            transition: opacity 0.25s ease;
        }
        #tr-view-scrape-overlay.tr-peek-mode {
            pointer-events: none;
            opacity: 0.12;
        }
        .tr-vpop-modal {
            background: #FFFFFF;
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 28px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.16), 0 4px 12px rgba(0, 0, 0, 0.05);
            width: 100%;
            max-width: 320px;
            color: #1A1C1E;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            animation: tr-modal-in 0.3s cubic-bezier(0.2, 0, 0, 1);
            position: relative;
        }
        .tr-vpop-book-area {
            padding: 32px 24px 26px;
            display: flex;
            flex-direction: column;
            align-items: center;
            position: relative;
        }
        .tr-vpop-book-wrap {
            width: 72px;
            height: 72px;
            border-radius: 20px;
            background: linear-gradient(135deg, #E8F0FE 0%, #D3E3FD 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 4px auto 18px;
            box-shadow: 0 8px 24px rgba(11, 87, 208, 0.14);
            position: relative;
        }
        .tr-vpop-spinner-ring {
            position: absolute;
            inset: -12px;
            border-radius: 50%;
            border: 2.5px solid rgba(11, 87, 208, 0.12);
            border-top: 3px solid #0B57D0;
            border-right: 3px solid #388af6;
            border-bottom: 3px solid transparent;
            pointer-events: none;
            animation: tr-vpop-ring-spin 1.8s linear infinite;
            filter: drop-shadow(0 2px 6px rgba(11, 87, 208, 0.22));
        }
        @keyframes tr-vpop-ring-spin {
            0% {
                transform: rotate(0deg);
            }
            100% {
                transform: rotate(360deg);
            }
        }
        .tr-vpop-book-svg {
            transform-origin: center center;
        }
        .tr-vpop-counter-wrap {
            display: flex;
            flex-direction: column;
            align-items: center;
            position: relative;
        }
        .tr-vpop-counter-label {
            font-size: 11.5px;
            font-weight: 700;
            letter-spacing: 0.6px;
            text-transform: uppercase;
            color: #5E6368;
            margin-bottom: 2px;
        }
        .tr-vpop-counter-box {
            position: relative;
            display: inline-flex;
            align-items: baseline;
            justify-content: center;
            gap: 6px;
            min-height: 54px;
        }
        .tr-vpop-counter-num {
            font-size: 52px;
            font-weight: 800;
            line-height: 1;
            font-variant-numeric: tabular-nums;
            color: #0B57D0;
            transition: transform 0.25s cubic-bezier(0.175, 0.885, 0.32, 1.275), color 0.2s ease;
        }
        .tr-vpop-counter-num.tr-bump {
            transform: scale(1.22);
            color: #16A34A !important;
        }
        .tr-vpop-counter-unit {
            font-size: 15px;
            font-weight: 600;
            color: #5E6368;
        }
        .tr-vpop-plus-one {
            position: absolute;
            top: -14px;
            right: -34px;
            color: #16A34A;
            font-size: 28px;
            font-weight: 900;
            line-height: 1;
            pointer-events: none;
            z-index: 10;
            animation: tr-plus-float 1.25s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        }
        @keyframes tr-plus-float {
            0% {
                opacity: 0;
                transform: translateY(12px) scale(0.6);
                filter: drop-shadow(0 0 2px rgba(22, 163, 74, 0.4));
            }
            25% {
                opacity: 1;
                transform: translateY(-2px) scale(1.3);
                filter: drop-shadow(0 0 12px rgba(22, 163, 74, 0.8));
            }
            60% {
                opacity: 1;
                transform: translateY(-20px) scale(1.1);
                filter: drop-shadow(0 0 10px rgba(22, 163, 74, 0.6));
            }
            100% {
                opacity: 0;
                transform: translateY(-38px) scale(0.9);
                filter: drop-shadow(0 0 0 transparent);
            }
        }
        .tr-vpop-status-card {
            width: 100%;
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 18px;
            padding: 12px 16px;
            margin-top: 14px;
            box-sizing: border-box;
            text-align: center;
        }
        .tr-vpop-book-title {
            font-size: 13.5px;
            font-weight: 600;
            color: #1E293B;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .tr-vpop-status-text {
            font-size: 12px;
            color: #0B57D0;
            font-weight: 500;
            margin-top: 4px;
        }
        .tr-vpop-bottom {
            display: none !important;
            padding: 16px 22px 20px;
            background: #FAFAFA;
            border-top: 1px solid #F1F5F9;
            flex-direction: column;
            gap: 8px;
        }
        .tr-vpop-bottom .tr-stepper-track-wrap {
            height: 8px;
            background: #E2E8F0;
            border-radius: 100px;
            overflow: hidden;
        }
        .tr-vpop-bottom .tr-stepper-active-line {
            height: 100%;
            background: #0B57D0;
            border-radius: 100px;
            transition: width 0.35s cubic-bezier(0.2, 0, 0, 1);
        }
        .tr-vpop-bottom .tr-stepper-meta {
            display: flex;
            justify-content: flex-end;
        }
        .tr-vpop-bottom .tr-stepper-tag {
            color: #0B57D0;
            font-size: 11.5px;
            font-weight: 700;
            font-variant-numeric: tabular-nums;
            text-shadow: none;
        }
        .tr-vpop-actions {
            display: flex;
            gap: 10px;
            margin-top: 6px;
        }
        .tr-vpop-btn-skip, .tr-vpop-btn-stop {
            flex: 1;
            padding: 10px 18px;
            border-radius: 100px;
            font-family: inherit;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            border: none;
        }
        .tr-vpop-btn-skip {
            background: #FFFFFF;
            border: 1px solid #CBD5E1 !important;
            color: #334155;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        }
        .tr-vpop-btn-skip:hover:not(:disabled) {
            background: #F1F5F9;
            color: #0F172A;
        }
        .tr-vpop-btn-stop {
            background: #FFDAD6;
            border: 1px solid rgba(186, 26, 26, 0.25) !important;
            color: #BA1A1A;
        }
        .tr-vpop-btn-stop:hover:not(:disabled) {
            background: #FFCDD2;
            color: #990000;
        }

        /* ==========================================================================
        MATERIAL 3 DIRECTION DIALOG & START BOOK GRID DIALOG
        ========================================================================== */
        #tr-direction-overlay, #tr-grid-start-overlay {
            position: fixed; inset: 0; background: rgba(0,0,0,0.55);
            backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px);
            z-index: 2147483647; display: flex; align-items: center; justify-content: center;
            font-family: 'Roboto', -apple-system, 'Segoe UI', sans-serif;
            animation: tr-cd-fade-in .2s ease-out;
        }
        .tr-direction-box {
            background: #FFFFFF;
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 28px; width: 92vw; max-width: 480px;
            box-shadow: 0 24px 60px rgba(0,0,0,0.22), 0 4px 16px rgba(0,0,0,0.06);
            animation: tr-modal-in .3s cubic-bezier(.2,0,0,1);
            overflow: hidden; display: flex; flex-direction: column;
        }
        .tr-direction-header {
            padding: 24px 24px 16px; text-align: center;
        }
        .tr-direction-icon {
            width: 56px; height: 56px; border-radius: 50%;
            background: #E8F0FE; display: flex; align-items: center; justify-content: center;
            margin: 0 auto 12px;
        }
        .tr-direction-title {
            font-size: 19px; font-weight: 700; color: #1A1C1E;
        }
        .tr-direction-sub {
            font-size: 13px; color: #64748B; margin-top: 5px; line-height: 1.45;
        }
        .tr-direction-options {
            padding: 6px 18px 14px; display: flex; flex-direction: column; gap: 10px;
        }
        .tr-direction-card {
            display: flex; align-items: center; gap: 14px; padding: 14px 16px;
            border-radius: 18px; background: #F8FAFD;
            border: 1.5px solid #E2E8F0; cursor: pointer;
            transition: all .2s cubic-bezier(0.2, 0, 0, 1);
        }
        .tr-direction-card:hover {
            background: #FFFFFF; border-color: #0B57D0;
            transform: translateY(-2px);
            box-shadow: 0 8px 24px rgba(11,87,208,0.14);
        }
        .tr-direction-card:active {
            transform: scale(0.98);
        }
        .tr-direction-card-icon {
            flex-shrink: 0; width: 40px; height: 40px; border-radius: 12px;
            background: #F1F5F9; color: #475569;
            display: flex; align-items: center; justify-content: center;
            transition: all .2s cubic-bezier(0.2, 0, 0, 1);
        }
        .tr-direction-card-icon.tr-dir-primary {
            background: #E8F0FE; color: #0B57D0;
        }
        .tr-direction-card:hover .tr-direction-card-icon {
            background: #E8F0FE; color: #0B57D0;
        }
        .tr-direction-card-body {
            flex: 1; min-width: 0;
        }
        .tr-direction-card-title-row {
            display: flex; align-items: center; gap: 8px; margin-bottom: 3px;
        }
        .tr-direction-card-title {
            font-size: 14.5px; font-weight: 700; color: #1E293B;
        }
        .tr-direction-badge {
            font-size: 10.5px; font-weight: 700; color: #0B57D0;
            background: #E8F0FE; padding: 2px 8px; border-radius: 100px;
            letter-spacing: 0.3px;
        }
        .tr-direction-badge-sec {
            font-size: 10.5px; font-weight: 700; color: #475569;
            background: #F1F5F9; padding: 2px 8px; border-radius: 100px;
        }
        .tr-direction-card-desc {
            font-size: 12px; color: #64748B; line-height: 1.4;
        }
        .tr-direction-card-arrow {
            color: #94A3B8; flex-shrink: 0;
            display: flex; align-items: center; justify-content: center;
            transition: transform .2s, color .2s;
        }
        .tr-direction-card:hover .tr-direction-card-arrow {
            transform: translateX(3px); color: #0B57D0;
        }
        .tr-direction-footer {
            padding: 12px 24px 16px; border-top: 1px solid #F1F5F9;
            display: flex; justify-content: flex-end; background: #FAFAFA;
        }

        /* Start Book Grid Dialog */
        .tr-grid-start-box {
            background: #FFFFFF;
            border: 1px solid rgba(226, 232, 240, 0.9);
            border-radius: 28px; width: 92vw; max-width: 660px; max-height: 88vh;
            box-shadow: 0 24px 60px rgba(0,0,0,0.24), 0 4px 16px rgba(0,0,0,0.06);
            animation: tr-modal-in .3s cubic-bezier(.2,0,0,1);
            overflow: hidden; display: flex; flex-direction: column;
        }
        .tr-grid-start-header {
            padding: 20px 22px 14px; display: flex; align-items: center; gap: 14px;
            border-bottom: 1px solid #F1F5F9; background: #FAFAFA;
        }
        .tr-grid-start-icon {
            width: 44px; height: 44px; border-radius: 14px;
            background: #E8F0FE; display: flex; align-items: center; justify-content: center;
            flex-shrink: 0;
        }
        .tr-grid-start-title {
            font-size: 17px; font-weight: 700; color: #1E293B;
        }
        .tr-grid-start-sub {
            font-size: 12px; color: #64748B; margin-top: 2px;
        }
        .tr-grid-start-search-wrap {
            padding: 12px 20px 8px; background: #FFFFFF;
        }
        .tr-grid-start-search-wrap input {
            width: 100%; padding: 10px 16px; border-radius: 100px;
            border: 1.5px solid #CBD5E1; font-family: inherit; font-size: 13px;
            outline: none; background: #F8FAFD; box-sizing: border-box;
            transition: border-color .15s, box-shadow .15s, background .15s;
        }
        .tr-grid-start-search-wrap input:focus {
            border-color: #0B57D0; background: #FFFFFF;
            box-shadow: 0 0 0 3px rgba(11,87,208,0.15);
        }
        .tr-grid-start-container {
            flex: 1; overflow-y: auto; padding: 12px 20px;
            display: grid; grid-template-columns: repeat(auto-fill, minmax(105px, 1fr));
            gap: 12px; min-height: 220px; max-height: 420px;
        }
        .tr-book-square {
            aspect-ratio: 1; border-radius: 18px; background: #F8FAFD;
            border: 1.5px solid #E2E8F0; display: flex; flex-direction: column;
            align-items: center; justify-content: center; padding: 8px 6px;
            cursor: pointer; position: relative; transition: all .2s cubic-bezier(0.2, 0, 0, 1);
            text-align: center; user-select: none; box-sizing: border-box;
        }
        .tr-book-square:hover {
            background: #FFFFFF; border-color: #0B57D0;
            transform: translateY(-3px) scale(1.02);
            box-shadow: 0 8px 20px rgba(11,87,208,0.18);
        }
        .tr-book-square:active {
            transform: scale(0.96);
        }
        .tr-book-square-num {
            position: absolute; top: 6px; left: 8px;
            font-size: 10px; font-weight: 700; color: #0B57D0;
            background: #E8F0FE; padding: 1px 6px; border-radius: 100px;
        }
        .tr-book-square-icon {
            margin: 6px 0 4px; display: flex; align-items: center; justify-content: center;
        }
        .tr-book-square-title {
            font-size: 11px; font-weight: 600; color: #1E293B;
            line-height: 1.25; display: -webkit-box; -webkit-line-clamp: 2;
            -webkit-box-orient: vertical; overflow: hidden; text-overflow: ellipsis;
            word-break: break-word; padding: 0 4px;
        }
        .tr-grid-start-footer {
            padding: 12px 20px 16px; border-top: 1px solid #F1F5F9;
            display: flex; align-items: center; justify-content: space-between;
            background: #FAFAFA; gap: 12px;
        }
        .tr-grid-start-preview {
            flex: 1; font-size: 12px; color: #475569; font-weight: 500;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        }
        .tr-grid-start-actions {
            display: flex; gap: 8px; flex-shrink: 0;
        }
        `;
        document.head.appendChild(style);
    };


    const ICONS = {
        book: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>',
        heart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"/></svg>',
        stop: '<svg viewBox="0 0 24 24" fill="currentColor"><rect x="6" y="6" width="12" height="12" rx="2"/></svg>',
        skip: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 5v14l9-7z"/><path d="M19 5v14"/></svg>',
        chevron: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>',
        zap: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" fill="currentColor"/></svg>',
        clock: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
        hourglass: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2h12M6 22h12M6 2v6l6 4-6 4v6M18 2v6l-6 4 6 4v6"/></svg>',
        flag: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 22V4M4 4h13l-2 5 2 5H4"/></svg>',
        search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
        check: '<svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'
    };

    const LOGO_DATA_URL = "data:image/png;base64,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";

    const createUI = () => {
        injectStyles();
        const panel = document.createElement('div');
        panel.className = 'tr-panel';
        const header = document.createElement('div');
        header.className = 'tr-header';
        header.innerHTML = `
                    <img class="tr-logo" src="${LOGO_DATA_URL}" alt="Logo AutoEbook">
                    <div class="tr-title">Auto Ebook<span class="tr-subtitle">By Em Wizi With Love <3</span></div>`;

        const collapseBtn = document.createElement('button');
        collapseBtn.className = 'tr-collapse';
        collapseBtn.title = 'Thu gọn';
        collapseBtn.innerHTML = ICONS.chevron;
        header.appendChild(collapseBtn);

        const body = document.createElement('div');
        body.className = 'tr-body';

        const etaWrap = document.createElement('div');
        etaWrap.className = 'tr-eta';
        etaWrap.innerHTML = `
                    <div class="tr-eta-row eta-main">
                        <span class="tr-eta-label">${ICONS.hourglass}<span>Sách hiện tại</span></span>
                        <span class="tr-eta-value" id="eta-current">--</span>
                    </div>
                    <div class="tr-eta-row">
                        <span class="tr-eta-label">${ICONS.flag}<span>Còn lại cả danh sách</span></span>
                        <span class="tr-eta-value" id="eta-total">--</span>
                    </div>
                    <div class="tr-eta-row">
                        <span class="tr-eta-label">${ICONS.clock}<span>Dự kiến xong lúc</span></span>
                        <span class="tr-eta-value" id="eta-finish">--</span>
                    </div>
                    <div class="tr-eta-row">
                        <span class="tr-eta-label">${ICONS.clock}<span>Tốc độ TB (thực tế)</span></span>
                        <span class="tr-eta-value" id="eta-speed">--</span>
                    </div>`;

        applyAnimOff();

        const sliderStepper = document.createElement('div');
        sliderStepper.className = 'tr-slider-stepper';
        sliderStepper.innerHTML = `
                    <div class="tr-stepper-track-wrap">
                        <div class="tr-stepper-active-line" id="tr-stepper-fill"></div>
                    </div>
                    <div class="tr-stepper-meta"><span class="tr-stepper-tag" id="tr-stepper-tag">0 / 0 cuốn · 0%</span></div>
                `;

        const btnAutoAll = document.createElement('button');
        btnAutoAll.className = 'tr-btn tr-btn-primary';
        btnAutoAll.innerHTML = ICONS.zap + '<span>Bắt đầu</span>';

        const btnScan = document.createElement('button');
        btnScan.style.display = 'none';

        const btnRead = document.createElement('button');
        btnRead.className = 'tr-btn tr-btn-secondary';
        btnRead.innerHTML = ICONS.heart + '<span>Đọc &amp; Like Trang Này</span>';

        const btnSkip = document.createElement('button');
        btnSkip.className = 'tr-btn tr-btn-secondary';
        btnSkip.style.display = 'none';
        btnSkip.innerHTML = ICONS.skip + '<span>Bỏ qua cuốn này</span>';

        const btnStop = document.createElement('button');
        btnStop.className = 'tr-btn tr-btn-danger';
        btnStop.style.display = 'none';
        btnStop.innerHTML = ICONS.stop + '<span>Dừng Tự Động</span>';

        const btnClearHist = document.createElement('button');
        btnClearHist.className = 'tr-btn tr-btn-outline-danger';
        btnClearHist.innerHTML = '<svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg><span>Xóa Lịch Sử Đọc</span>';
        btnClearHist.addEventListener('click', async () => {
            const curHist = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
            if (curHist.length === 0) {
                alert('Lịch sử đọc hiện đang trống!');
                return;
            }
            const ok = await showConfirmDialog(
                'Xóa lịch sử sách đã đọc',
                `Bạn có chắc muốn xóa lịch sử <strong>${curHist.length} cuốn sách</strong> đã đọc?<br><br>Sau khi xóa, auto sẽ đọc lại toàn bộ các cuốn sách này từ đầu.`,
                'Xóa lịch sử',
                'Huỷ'
            );
            if (ok) {
                localStorage.removeItem(KEY_READ_HISTORY);
                clearBookCache();
                alert(' Đã xóa toàn bộ lịch sử sách đã đọc!');
                updateStatus(_d('IMSQw6MgeMOzYSBs4buLY2ggc+G7rSDEkeG7jWMg4oCUIHPhurVuIHPDoG5nIMSR4buNYyBs4bqhaSB04burIMSR4bqndQ=='), 'success');
            }
        });

        const btnGroup = document.createElement('div');
        btnGroup.className = 'tr-btn-group';
        btnGroup.appendChild(btnAutoAll);
        btnGroup.appendChild(btnSkip);
        btnGroup.appendChild(btnStop);

        const btnSecRow = document.createElement('div');
        btnSecRow.className = 'tr-btn-row';
        btnSecRow.appendChild(btnRead);
        btnSecRow.appendChild(btnClearHist);
        btnGroup.appendChild(btnSecRow);

        const viewCountDebug = document.createElement('section');
        viewCountDebug.className = 'tr-view-count-debug';
        viewCountDebug.style.display = 'none';

        body.appendChild(btnGroup);
        body.appendChild(sliderStepper);
        const footerNote = document.createElement('div');
        footerNote.className = 'tr-footer-note';
        footerNote.textContent = getFooterNoteText();
        body.appendChild(footerNote);

        panel.appendChild(header);
        panel.appendChild(body);
        document.body.appendChild(panel);

        if (localStorage.getItem(KEY_UI_COLLAPSED) === 'true') panel.classList.add('tr-collapsed');
        collapseBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const c = panel.classList.toggle('tr-collapsed');
            localStorage.setItem(KEY_UI_COLLAPSED, c ? 'true' : 'false');
        });
        header.addEventListener('click', () => {
            const collapsed = panel.classList.toggle('tr-collapsed');
            localStorage.setItem(KEY_UI_COLLAPSED, collapsed ? 'true' : 'false');
        });

        return {
            btnRead, btnAutoAll, btnStop, btnSkip, btnScan,
            etaWrap,
            etaCurrent: etaWrap.querySelector('#eta-current'),
            etaTotal: etaWrap.querySelector('#eta-total'),
            etaFinish: etaWrap.querySelector('#eta-finish'),
            etaSpeed: etaWrap.querySelector('#eta-speed'),
            stepperWrap: sliderStepper,
            stepperFill: sliderStepper.querySelector('#tr-stepper-fill'),
            stepperTag: sliderStepper.querySelector('#tr-stepper-tag'),
            stepperText: null,
            viewCountDebug,
            viewCountDebugBook: viewCountDebug.querySelector('#tr-vcd-book'),
            viewCountDebugBefore: viewCountDebug.querySelector('#tr-vcd-before'),
            viewCountDebugAfter: viewCountDebug.querySelector('#tr-vcd-after'),
            viewCountDebugObserved: viewCountDebug.querySelector('#tr-vcd-observed'),
            viewCountDebugStatus: viewCountDebug.querySelector('#tr-vcd-status'),
            panel
        };
    };

    const createOrShowViewScrapePopup = (ui) => {
        let overlay = document.getElementById('tr-view-scrape-overlay');
        const loopTotal = parseInt(localStorage.getItem(KEY_VIEW_LOOP_TOTAL) || '1', 10);
        const loopCurrent = parseInt(localStorage.getItem(KEY_VIEW_LOOP_CURRENT) || '1', 10);
        const loopText = loopTotal > 1 ? `Vòng ${loopCurrent}/${loopTotal}` : 'Vòng 1/1';
        const successCount = parseInt(localStorage.getItem(KEY_READ_SUCCESS_COUNT) || '0', 10);

        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'tr-view-scrape-overlay';
            overlay.innerHTML = `
                    <div class="tr-vpop-modal">
                        <div class="tr-vpop-book-area">
                            <div class="tr-vpop-book-wrap">
                                <div class="tr-vpop-spinner-ring"></div>
                                <svg viewBox="0 0 48 48" width="38" height="38" fill="none" class="tr-vpop-book-svg">
                                    <path d="M8 38V10C8 8.89543 8.89543 8 10 8H20C22.2091 8 24 9.79086 24 12V40C24 38.8954 23.1046 38 22 38H10C8.89543 38 8 38 8 38Z" fill="#0B57D0"/>
                                    <path d="M40 38V10C40 8.89543 39.1046 8 38 8H28C25.7909 8 24 9.79086 24 12V40C24 38.8954 24.8954 38 26 38H38C39.1046 38 40 38 40 38Z" fill="#1A73E8"/>
                                    <path d="M13 15H19M13 21H19M13 27H19M13 33H19" stroke="#FFFFFF" stroke-width="2" stroke-linecap="round"/>
                                    <path d="M29 15H35M29 21H35M29 27H35M29 33H35" stroke="#E8F0FE" stroke-width="2" stroke-linecap="round"/>
                                    <path d="M22 8V24L24 22L26 24V8" fill="#FBBC04"/>
                                </svg>
                            </div>

                            <div class="tr-vpop-counter-wrap">
                                <div class="tr-vpop-counter-box" id="tr-vpop-counter-box">
                                    <span class="tr-vpop-counter-num" id="tr-vpop-counter-num">${successCount}</span>
                                    <span class="tr-vpop-counter-unit">cuốn</span>
                                </div>
                            </div>
                        </div>

                        <div class="tr-vpop-bottom" style="display: none !important;">
                            <div class="tr-stepper-track-wrap">
                                <div class="tr-stepper-active-line" id="tr-vpop-stepper-fill"></div>
                            </div>
                            <div class="tr-stepper-meta">
                                <span class="tr-vpop-stepper-tag" id="tr-vpop-stepper-tag">0 / 0 cuốn · 0%</span>
                            </div>
                            <div class="tr-vpop-actions">
                                <button type="button" class="tr-vpop-btn-skip" id="tr-vpop-btn-skip">Bỏ qua cuốn này</button>
                                <button type="button" class="tr-vpop-btn-stop" id="tr-vpop-btn-stop">Dừng tự động</button>
                            </div>
                        </div>
                    </div>
                `;
            document.body.appendChild(overlay);

            const skipBtn = overlay.querySelector('#tr-vpop-btn-skip');
            if (skipBtn) {
                skipBtn.addEventListener('click', () => {
                    if (ui && ui.btnSkip) ui.btnSkip.click();
                });
            }

            const stopBtn = overlay.querySelector('#tr-vpop-btn-stop');
            if (stopBtn) {
                stopBtn.addEventListener('click', () => {
                    if (ui && ui.btnStop) ui.btnStop.click();
                });
            }
        } else {
            const numEl = overlay.querySelector('#tr-vpop-counter-num');
            if (numEl) numEl.textContent = successCount.toString();
        }
        return overlay;
    };

    const closeViewScrapePopup = () => {
        const overlay = document.getElementById('tr-view-scrape-overlay');
        if (overlay) overlay.remove();
    };

    const updateViewPopupProgress = (index, total, pct, loopCurrent, loopTotal) => {
        const overlay = document.getElementById('tr-view-scrape-overlay');
        if (!overlay) return;
        const fill = overlay.querySelector('#tr-vpop-stepper-fill');
        if (fill) fill.style.width = pct + '%';
        const tag = overlay.querySelector('#tr-vpop-stepper-tag');
        const loopText = loopTotal > 1 ? ` · Vòng ${loopCurrent}/${loopTotal}` : '';
        if (tag) tag.innerText = index + ' / ' + total + ' cuốn · ' + pct + '%' + loopText;
        const loopBadge = overlay.querySelector('#tr-vpop-loop-badge');
        if (loopBadge && loopTotal > 1) loopBadge.textContent = `Vòng ${loopCurrent}/${loopTotal}`;
    };

    const updateViewPopupStatus = (statusText) => {
        const overlay = document.getElementById('tr-view-scrape-overlay');
        if (!overlay || !statusText) return;
        const stEl = overlay.querySelector('#tr-vpop-status-text');
        if (stEl) stEl.textContent = statusText;
    };

    const updateViewPopupBookTitle = (bookTitle) => {
        const overlay = document.getElementById('tr-view-scrape-overlay');
        if (!overlay || !bookTitle) return;
        const cleanTitle = String(bookTitle).replace(/\s*\d*\s*lượt\s*xem.*$/i, '').trim();
        const btEl = overlay.querySelector('#tr-vpop-book-title');
        if (btEl) btEl.textContent = '📖 ' + (cleanTitle || bookTitle);
    };

    const showViewPopupPlusOne = (newCount) => {
        const overlay = document.getElementById('tr-view-scrape-overlay');
        if (!overlay) return;
        const box = overlay.querySelector('#tr-vpop-counter-box');
        const numEl = overlay.querySelector('#tr-vpop-counter-num');
        if (numEl) {
            numEl.textContent = newCount.toString();
            numEl.classList.remove('tr-bump');
            void numEl.offsetWidth;
            numEl.classList.add('tr-bump');
            setTimeout(() => numEl.classList.remove('tr-bump'), 450);
        }
        if (box) {
            const badge = document.createElement('div');
            badge.className = 'tr-vpop-plus-one';
            badge.textContent = '+1';
            box.appendChild(badge);
            setTimeout(() => badge.remove(), 1300);
        }
    };

    const updateProgress = (ui, index, total, detailText = '') => {
        const pct = total > 0 ? Math.min(100, Math.round((index / total) * 100)) : 0;
        if (ui.stepperFill) ui.stepperFill.style.width = pct + '%';

        const isViewMode = localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true';
        const loopTotal = parseInt(localStorage.getItem(KEY_VIEW_LOOP_TOTAL) || '1', 10);
        const loopCurrent = parseInt(localStorage.getItem(KEY_VIEW_LOOP_CURRENT) || '1', 10);
        const loopText = (isViewMode && loopTotal > 1) ? ` · Vòng ${loopCurrent}/${loopTotal}` : '';

        if (ui.stepperTag) ui.stepperTag.innerText = index + ' / ' + total + ' cuốn · ' + pct + '%' + loopText;
        if (detailText && ui.stepperText) ui.stepperText.innerText = detailText;

        if (isViewMode) {
            updateViewPopupProgress(index, total, pct, loopCurrent, loopTotal);
        }
    };

    const updateETA = (ui, index, total, currentStageMs = 0) => {
        return;
    };

    const enablePageBlocker = () => {
        let blocker = document.getElementById('tr-page-blocker');
        if (!blocker) {
            blocker = document.createElement('div');
            blocker.id = 'tr-page-blocker';
            ['click', 'mousedown', 'mouseup', 'touchstart', 'touchend', 'contextmenu'].forEach(evt => {
                blocker.addEventListener(evt, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                }, true);
            });
            document.body.appendChild(blocker);
        }
        blocker.style.display = 'block';
    };

    const disablePageBlocker = () => {
        const blocker = document.getElementById('tr-page-blocker');
        if (blocker) blocker.style.display = 'none';
    };

    const setNormalMode = (ui) => {
        disablePageBlocker();
        closeViewScrapePopup();
        ui.btnRead.style.display = 'flex';
        ui.btnAutoAll.style.display = 'flex';
        ui.btnScan.style.display = 'flex';
        ui.btnStop.style.display = 'none';
        ui.btnSkip.style.display = 'none';
        ui.btnRead.disabled = false;
        ui.btnAutoAll.disabled = false;
        ui.btnScan.disabled = false;
        ui.btnStop.disabled = false;
        ui.btnSkip.disabled = false;
        if (ui.viewCountDebug) ui.viewCountDebug.style.display = 'none';
        if (ui.panel) ui.panel.classList.remove('tr-panel-expanded');
    };

    const setRunningMode = (ui, canSkipBook = false) => {
        enablePageBlocker();
        ui.btnRead.style.display = 'none';
        ui.btnAutoAll.style.display = 'none';
        ui.btnScan.style.display = 'none';
        ui.btnStop.style.display = 'flex';
        ui.btnSkip.style.display = canSkipBook ? 'flex' : 'none';
        ui.btnStop.disabled = false;
        ui.btnSkip.disabled = false;
    };

    const handleFileAccessErrorIfNeeded = (ui, updateStatus, source) => {
        return false;
    };

    const initApp = () => {
        const ui = createUI();

        testImage(getBgUrl()).then((ok) => {
            if (!ok) {
                ui.panel.style.backgroundImage = 'linear-gradient(145deg, #F7F9FC, #EEF2F7)';
                ui.panel.style.backgroundSize = 'auto';
            }
        });

        const updateStatus = (text, type = 'info') => {
            const raw = typeof text === 'string' ? _d(text) : text;
            const cleanText = String(raw ?? '').replace(/^\s*\[AutoEbooks\]\s*/, '');
            const taggedText = `${APP_TAG} ${cleanText}`;

            // Chỉ hiện trong Console các thông tin thực sự cần thiết (Lỗi, Cảnh báo, Kết quả quan trọng)
            if (type === 'error') {
                console.error(taggedText);
            } else if (type === 'warn') {
                console.warn(taggedText);
            } else if (type === 'success') {
                // Lọc bỏ các bước phụ (tải bìa, đóng sách) để console luôn gọn gàng
                const isMinor = /tải xong bìa|đóng sách/i.test(cleanText);
                if (!isMinor) {
                    console.log(taggedText);
                }
            }

            if (ui.stepperText) ui.stepperText.innerText = taggedText;
            if (localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true') {
                updateViewPopupStatus(cleanText);
            }
            if (type === 'error') {
                showErrorAlert(taggedText);
            }
        };

        const updateViewCountDebug = (patch = {}) => {
            if (!ui.viewCountDebug) return;
            ui.viewCountDebug.style.display = 'none';
        };

        if (localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true') {
            updateViewCountDebug({ status: 'Đang khôi phục phiên cào lượt đọc...' });
        }

        const onCancelScrape = (partialBooks) => {
            setTimeout(() => {
                localStorage.removeItem(KEY_STOP);
                setNormalMode(ui);
                updateStatus(_d('IMSQw6MgaHXhu7cgY8OgbyDigJQgY8OybiA=') + partialBooks.length + _d('IHPDoWNo'), 'warn');
                if (partialBooks.length === 0) { updateStatus(_d('S2jDtG5nIGPDsyBzw6FjaCBuw6BvIGPDoG8gxJHGsOG7o2M='), 'warn'); return; }

                showBookPicker(partialBooks, ' Đã huỷ cào — ' + partialBooks.length + ' sách đã cào').then(pickResult => {
                    if (pickResult === null) { updateStatus(_d('xJDDoyBodeG7tyBjaOG7jW4gc8OhY2g='), 'warn'); return; }
                    let bookList = partialBooks.filter(b => pickResult.selectedUrls.has(b.url));
                    if (pickResult.limit > 0 && bookList.length > pickResult.limit) bookList = bookList.slice(0, pickResult.limit);
                    if (bookList.length === 0) { updateStatus(_d('IEtow7RuZyBjw7Mgc8OhY2ggbsOgbyDEkcaw4bujYyBjaOG7jW4='), 'warn'); return; }

                    localStorage.setItem(KEY_LIST, JSON.stringify(bookList));
                    localStorage.setItem(KEY_INDEX, '0');
                    localStorage.setItem(KEY_ACTIVE, 'true');
                    localStorage.setItem(KEY_START_TIME, Date.now().toString());
                    localStorage.removeItem(KEY_STOP);

                    const etaSec = bookList.length * getAvgSecPerBook();
                    updateProgress(ui, 0, bookList.length, ' Sẵn sàng ' + bookList.length + ' cuốn — dự kiến ' + fmtTime(etaSec));
                    setTimeout(() => { window.location.href = bookList[0].url; }, 1500);
                });
            }, 300);
        };

        if (isStopped()) {
            localStorage.removeItem(KEY_STOP);
            localStorage.removeItem(KEY_SKIP_CURRENT);
            localStorage.removeItem(KEY_ACTIVE);
            localStorage.removeItem(KEY_VIEW_COUNT_MODE);
            localStorage.removeItem(KEY_VIEW_LOOP_TOTAL);
            localStorage.removeItem(KEY_VIEW_LOOP_CURRENT);
            localStorage.removeItem(KEY_START_TIME);
            localStorage.removeItem(KEY_SCRAPE_PENDING);
            localStorage.removeItem(KEY_LIST);
            localStorage.removeItem(KEY_INDEX);
            localStorage.removeItem(KEY_SCAN_ACTIVE);
            localStorage.removeItem(KEY_SCAN_LIST);
            localStorage.removeItem(KEY_SCAN_INDEX);
        }

        const isScanning = localStorage.getItem(KEY_SCAN_ACTIVE) === 'true';

        if (isScanning) {
            setRunningMode(ui);
            const scanList = JSON.parse(localStorage.getItem(KEY_SCAN_LIST) || '[]');
            const scanIdx = parseInt(localStorage.getItem(KEY_SCAN_INDEX) || '0');
            updateProgress(ui, scanIdx, scanList.length, 'Đang quét danh sách...');
            if (isListPage() && scanList[scanIdx]) { window.location.href = scanList[scanIdx].url; return; }
            setTimeout(async () => {
                try { await runScanStep(ui, updateStatus, updateProgress); }
                catch (e) {
                    updateStatus(_d('IEzhu5dpIHNjYW46IA==') + e.message, 'error');
                    localStorage.removeItem(KEY_SCAN_ACTIVE);
                    setNormalMode(ui);
                }
            }, 1200);
        }

        const isActive = localStorage.getItem(KEY_ACTIVE) === 'true';

        if (!isActive && !isScanning && localStorage.getItem(KEY_SCRAPE_PENDING) === 'true') {
            localStorage.removeItem(KEY_SCRAPE_PENDING);
            if (isListPage()) setTimeout(() => ui.btnAutoAll.click(), 1500);
            else { window.location.href = LIST_URL; return; }
        }

        // Lúc nào vào web lần đầu (khi chưa chạy auto) cũng tự động tải lại data pack mới nhất từ GitHub
        if (!isActive && !isScanning) {
            setTimeout(() => {
                fetchBookPackFromGithub(updateStatus).catch(() => { });
            }, 600);
        }

        let list = JSON.parse(localStorage.getItem(KEY_LIST) || '[]');
        let index = parseInt(localStorage.getItem(KEY_INDEX) || '0');

        if (isActive && !isScanning) {
            if (!localStorage.getItem(KEY_START_TIME)) {
                localStorage.setItem(KEY_START_TIME, Date.now().toString());
            }
            setRunningMode(ui, true);
            ui.btnSkip.disabled = true;
            ui.panel.classList.add('tr-panel-expanded');
            if (localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true') {
                createOrShowViewScrapePopup(ui);
            }
            const targetLimitInit = parseInt(localStorage.getItem(KEY_TARGET_COUNT) || '0', 10);
            const currentSuccessInit = parseInt(localStorage.getItem(KEY_READ_SUCCESS_COUNT) || '0', 10);
            const totalForProgress = targetLimitInit > 0 ? targetLimitInit : list.length;
            const currentForProgress = targetLimitInit > 0 ? currentSuccessInit : index;
            updateProgress(ui, currentForProgress, totalForProgress);
            updateETA(ui, currentForProgress, totalForProgress, 0);

            if (isListPage()) {
                if (list[index]) { window.location.href = list[index].url; return; }
                else {
                    localStorage.removeItem(KEY_ACTIVE);
                    localStorage.removeItem(KEY_INDEX);
                    setNormalMode(ui);
                }
            }

            const runAutoProcess = async () => {
                if (isStopped()) return;
                ui.btnSkip.disabled = false;

                const targetLimit = parseInt(localStorage.getItem(KEY_TARGET_COUNT) || '0', 10);
                const isViewMode = localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true';
                let currentSuccess = parseInt(localStorage.getItem(KEY_READ_SUCCESS_COUNT) || '0', 10);

                const finishAll = async () => {
                    const loopTotal = parseInt(localStorage.getItem(KEY_VIEW_LOOP_TOTAL) || '1', 10);
                    const startTime = parseInt(localStorage.getItem(KEY_START_TIME) || '0', 10);
                    const durationSec = startTime > 0 ? Math.max(1, Math.round((Date.now() - startTime) / 1000)) : 0;
                    const durationText = durationSec > 0 ? fmtTime(durationSec) : '--';

                    localStorage.removeItem(KEY_ACTIVE);
                    localStorage.removeItem(KEY_LIMIT);
                    localStorage.removeItem(KEY_TARGET_COUNT);
                    localStorage.removeItem(KEY_VIEW_COUNT_MODE);
                    localStorage.removeItem(KEY_VIEW_LOOP_TOTAL);
                    localStorage.removeItem(KEY_VIEW_LOOP_CURRENT);
                    localStorage.removeItem(KEY_START_TIME);
                    setNormalMode(ui);

                    const finalRead = parseInt(localStorage.getItem(KEY_READ_SUCCESS_COUNT) || '0', 10);
                    const totalTarget = targetLimit > 0 ? targetLimit : list.length;
                    const loopNote = loopTotal > 1 ? ` (${loopTotal} vòng lặp)` : '';
                    updateProgress(ui, totalTarget, totalTarget, ' Đã hoàn tất! Đã đọc: ' + finalRead + ' cuốn' + loopNote);

                    await showSummaryDialog({
                        totalTarget: targetLimit,
                        successCount: finalRead,
                        durationText: durationText
                    });
                };

                if (targetLimit > 0 && currentSuccess >= targetLimit) {
                    await finishAll();
                    return;
                }

                if (index >= list.length) {
                    await finishAll();
                    return;
                }

                const currentBook = list[index];
                const bookStart = Date.now();
                const totalTargetForETA = targetLimit > 0 ? targetLimit : list.length;
                const currentForETA = targetLimit > 0 ? currentSuccess : index;
                const tickETA = setInterval(() => {
                    if (isStopped()) { clearInterval(tickETA); return; }
                    updateETA(ui, currentForETA, totalTargetForETA, Date.now() - bookStart);
                }, 500);

                const shortTitle = currentBook.title.length > 35 ? currentBook.title.slice(0, 35) + '...' : currentBook.title;
                const progText = targetLimit > 0
                    ? ' Mục tiêu: ' + currentSuccess + '/' + targetLimit + ' cuốn (Đang duyệt #' + (index + 1) + ') — "' + shortTitle + '"'
                    : ' Đang xử lý cuốn ' + (index + 1) + '/' + list.length + ' — "' + shortTitle + '"';
                updateProgress(ui, currentForETA, totalTargetForETA, progText);
                if (localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true') {
                    updateViewPopupBookTitle(currentBook.title);
                }

                const addSkippedReason = (reason) => {
                    let skippedList = JSON.parse(localStorage.getItem(KEY_SKIPPED_DETAILS) || '[]');
                    skippedList.push({
                        title: currentBook.title || 'Không rõ tên',
                        url: currentBook.url,
                        reason: reason
                    });
                    localStorage.setItem(KEY_SKIPPED_DETAILS, JSON.stringify(skippedList));
                };

                const addReserveReplacement = () => {
                    if (targetLimit <= 0) return false;
                    const reserve = JSON.parse(localStorage.getItem('TR_RESERVE_LIST') || '[]');
                    while (reserve.length > 0) {
                        const nextBook = reserve.shift();
                        if (!nextBook || !nextBook.url || list.some(book => book.url === nextBook.url)) continue;
                        list.push(nextBook);
                        localStorage.setItem(KEY_LIST, JSON.stringify(list));
                        localStorage.setItem('TR_RESERVE_LIST', JSON.stringify(reserve));
                        return true;
                    }
                    localStorage.setItem('TR_RESERVE_LIST', JSON.stringify(reserve));
                    return false;
                };

                let skipped = false;
                const consumeSkipRequest = () => {
                    if (localStorage.getItem(KEY_SKIP_CURRENT) !== 'true') return false;
                    localStorage.removeItem(KEY_SKIP_CURRENT);
                    skipped = true;
                    addSkippedReason('Bỏ qua theo yêu cầu');
                    addReserveReplacement();
                    updateStatus(_d('IMSQw6MgYuG7jyBxdWEgdGhlbyB5w6p1IGPhuqd1IOKAlCBjaHV54buDbiBzYW5nIGN14buRbiB0aeG6v3AgdGhlbw=='), 'warn');
                    return true;
                };

                const viewCountMode = localStorage.getItem(KEY_VIEW_COUNT_MODE) === 'true';
                if (viewCountMode) {
                    updateViewCountDebug({
                        title: currentBook.title,
                        before: null,
                        after: null,
                        observed: null,
                        status: 'Đang bắt đầu kiểm tra lượt xem...'
                    });
                }
                try {
                    const readResult = await performAutoReadLogic(updateStatus, viewCountMode
                        ? { forceRead: true, verifyViewCount: true, bookTitle: currentBook.title, onViewCountDebug: updateViewCountDebug }
                        : { bookTitle: currentBook.title });
                    let viewCountConfirmed = true;
                    if (viewCountMode) {
                        const before = readResult?.viewCountBefore;
                        const after = readResult?.viewCountAfter;
                        if (before === null || before === undefined || after === null || after === undefined) {
                            viewCountConfirmed = false;
                            skipped = true;
                            updateViewCountDebug({
                                before,
                                after,
                                observed: after ?? null,
                                status: 'Không đủ dữ liệu để kết luận lượt xem có tăng'
                            });
                            addSkippedReason('Không xác định được lượt xem trước/sau để kiểm tra');
                            updateStatus(_d('IEtow7RuZyB4w6FjIMSR4buLbmggxJHGsOG7o2MgbMaw4bujdCB4ZW0gdHLGsOG7m2Mvc2F1IGPhu6dhIGN14buRbiBuw6B5'), 'warn');
                        } else if (before >= 1000 && after >= before) {
                            // Sách có lượt xem tính bằng ngàn (K) — web làm tròn nên không hiện số lẻ, hoan hỉ coi như đã đọc
                            viewCountConfirmed = true;
                            updateViewCountDebug({
                                title: currentBook.title,
                                before,
                                after,
                                observed: after,
                                status: 'Lượt xem lớn dạng K (' + before + ') làm tròn không hiện số lẻ — hoan hỉ xác nhận đã đọc'
                            });
                            updateStatus(_d('IFPDoWNoIGPDsyBsxrDhu6N0IHhlbSBk4bqhbmcgSyAo') + before + _d('KSDigJQgaG9hbiBo4buJIHjDoWMgbmjhuq1uIMSRw6MgxJHhu41jIHRow6BuaCBjw7RuZyE='), 'success');
                        } else if (after <= before) {
                            viewCountConfirmed = false;
                            skipped = true;
                            updateViewCountDebug({
                                title: currentBook.title,
                                before,
                                after,
                                observed: after,
                                status: 'Lượt xem không tăng; có thể sách đã đọc hoặc số liệu cập nhật trễ'
                            });
                            addSkippedReason('Có thể sách này đã được đọc rồi — lượt xem không tăng (' + before + ' → ' + after + ')');
                            updateStatus(_d('IEPDsyB0aOG7gyBzw6FjaCBuw6B5IMSRw6MgxJHGsOG7o2MgxJHhu41jIHLhu5NpIOKAlCBsxrDhu6N0IHhlbSBraMO0bmcgdMSDbmcgKA==') + before + ' → ' + after + ')', 'warn');
                        } else {
                            updateViewCountDebug({
                                title: currentBook.title,
                                before,
                                after,
                                observed: after,
                                status: 'Lượt xem đã tăng; hệ thống xác nhận sách'
                            });
                            updateStatus(_d('IEzGsOG7o3QgeGVtIMSRw6MgdMSDbmc6IA==') + before + ' → ' + after, 'success');
                        }
                    }
                    let history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
                    if (!history.includes(currentBook.url)) {
                        history.push(currentBook.url);
                        localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
                    }
                    if (viewCountConfirmed) {
                        currentSuccess++;
                        localStorage.setItem(KEY_READ_SUCCESS_COUNT, currentSuccess.toString());
                        if (viewCountMode) {
                            showViewPopupPlusOne(currentSuccess);
                            await sleep(1200);
                        }
                    }
                }
                catch (err) {
                    if (viewCountMode) {
                        updateViewCountDebug({ title: currentBook.title, status: 'Đo lượt xem bị gián đoạn: ' + (err?.message || 'lỗi không rõ') });
                    }
                    if (consumeSkipRequest()) {
                    } else if (err.message === 'Đã dừng.') {
                        clearInterval(tickETA);
                        updateStatus(_d('IMSQw6MgZOG7q25nIHRoZW8gecOqdSBj4bqndQ=='), 'warn');
                        return;
                    }
                    else {
                        if (err.message === 'FILE_ACCESS_ERROR') {
                            clearInterval(tickETA);
                            addSkippedReason('Lỗi truy cập file tài liệu');
                            const handled = handleFileAccessErrorIfNeeded(ui, updateStatus, 'read');
                            if (handled) return;
                        }
                        if (err.message === 'TEACHER_BOOK') {
                            skipped = true;
                            const _tc = parseInt(localStorage.getItem(KEY_TEACHER_COUNT) || '0', 10);
                            localStorage.setItem(KEY_TEACHER_COUNT, (_tc + 1).toString());
                            addSkippedReason('Sách giáo viên');
                            updateStatus(_d('IFPDoWNoIGdpw6FvIHZpw6pu'), 'warn');
                            let history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
                            if (!history.includes(currentBook.url)) {
                                history.push(currentBook.url);
                                localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
                            }
                            if (targetLimit > 0) {
                                let reserve = JSON.parse(localStorage.getItem('TR_RESERVE_LIST') || '[]');
                                if (reserve.length > 0) {
                                    const nextBook = reserve.shift();
                                    list.push(nextBook);
                                    localStorage.setItem(KEY_LIST, JSON.stringify(list));
                                    localStorage.setItem('TR_RESERVE_LIST', JSON.stringify(reserve));
                                }
                            }
                        }
                        else if (err.message === 'ALREADY_READ') {
                            skipped = true;
                            const _sc = parseInt(localStorage.getItem(KEY_SKIP_COUNT) || '0', 10);
                            localStorage.setItem(KEY_SKIP_COUNT, (_sc + 1).toString());
                            addSkippedReason('Đã hoàn thành trước đó');
                            let history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
                            if (!history.includes(currentBook.url)) {
                                history.push(currentBook.url);
                                localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
                            }
                            if (targetLimit > 0) {
                                let reserve = JSON.parse(localStorage.getItem('TR_RESERVE_LIST') || '[]');
                                if (reserve.length > 0) {
                                    const nextBook = reserve.shift();
                                    list.push(nextBook);
                                    localStorage.setItem(KEY_LIST, JSON.stringify(list));
                                    localStorage.setItem('TR_RESERVE_LIST', JSON.stringify(reserve));
                                }
                            }
                            const nextIdx = index + 1;
                            localStorage.setItem(KEY_INDEX, nextIdx.toString());
                            if (nextIdx < list.length) {
                                updateStatus(_d('IFPDoWNoIMSRw6MgxJHhu41jIHLhu5NpIOKAlCBsxrDhu5t0IG5nYXkgc2FuZyBjdeG7kW4gIw==') + (nextIdx + 1) + '...', 'info');
                                window.location.href = list[nextIdx].url;
                                return;
                            } else {
                                if (isViewMode) {
                                    const loopTotal = parseInt(localStorage.getItem(KEY_VIEW_LOOP_TOTAL) || '1', 10);
                                    const loopCurrent = parseInt(localStorage.getItem(KEY_VIEW_LOOP_CURRENT) || '1', 10);
                                    if (loopCurrent < loopTotal) {
                                        const nextLoop = loopCurrent + 1;
                                        localStorage.setItem(KEY_VIEW_LOOP_CURRENT, nextLoop.toString());
                                        localStorage.setItem(KEY_INDEX, '0');
                                        updateStatus(_d('IEhvw6BuIHRow6BuaCB2w7JuZyA=') + loopCurrent + '/' + loopTotal + _d('ISBC4bqvdCDEkeG6p3UgdsOybmcg') + nextLoop + '/' + loopTotal + _d('Li4u'), 'success');
                                        await sleep(1500);
                                        window.location.href = list[0].url;
                                        return;
                                    }
                                }
                                await finishAll();
                                return;
                            }
                        }
                        else if (err.message && err.message.startsWith('CHƯA_ĐẠT')) {
                            skipped = true;
                            addSkippedReason(err.message);
                            addReserveReplacement();
                            updateStatus(_d('IMSQw6MgY+G7kSDEkeG7jWMgbmjGsG5nIGNoxrBhIMSR4bqhdDog') + err.message + _d('IOKAlCBi4buPIHF1YQ=='), 'warn');
                            await sleep(1200);
                        }
                        else {
                            skipped = true;
                            addSkippedReason('Lỗi: ' + (err.message || 'Không xác định'));
                            updateStatus(_d('IEzhu5dpIOG7nyBjdeG7kW4gbsOgeTog') + err.message + _d('IOKAlCBi4buPIHF1YQ=='), 'warn');
                            await sleep(1200);
                            addReserveReplacement();
                        }
                    }
                }

                consumeSkipRequest();

                clearInterval(tickETA);
                const elapsedMs = Date.now() - bookStart;
                if (!skipped) recordBookTime(elapsedMs);
                clearRetry(currentBook.url);

                if (isStopped()) { updateStatus(_d('IMSQw6MgZOG7q25nIHRoZW8gecOqdSBj4bqndQ=='), 'warn'); return; }

                const nextIdx = index + 1;
                localStorage.setItem(KEY_INDEX, nextIdx.toString());

                const isGoalReached = targetLimit > 0 && currentSuccess >= targetLimit;

                if (isGoalReached || nextIdx >= list.length) {
                    if (isViewMode && !isGoalReached) {
                        const loopTotal = parseInt(localStorage.getItem(KEY_VIEW_LOOP_TOTAL) || '1', 10);
                        const loopCurrent = parseInt(localStorage.getItem(KEY_VIEW_LOOP_CURRENT) || '1', 10);
                        if (loopCurrent < loopTotal) {
                            const nextLoop = loopCurrent + 1;
                            localStorage.setItem(KEY_VIEW_LOOP_CURRENT, nextLoop.toString());
                            localStorage.setItem(KEY_INDEX, '0');
                            updateStatus(_d('IEhvw6BuIHRow6BuaCB2w7JuZyA=') + loopCurrent + '/' + loopTotal + _d('ISBC4bqvdCDEkeG6p3UgdsOybmcg') + nextLoop + '/' + loopTotal + _d('Li4u'), 'success');
                            await sleep(1500);
                            window.location.href = list[0].url;
                            return;
                        }
                    }
                    await finishAll();
                    return;
                }

                const loopTotal = parseInt(localStorage.getItem(KEY_VIEW_LOOP_TOTAL) || '1', 10);
                const loopCurrent = parseInt(localStorage.getItem(KEY_VIEW_LOOP_CURRENT) || '1', 10);
                const loopPrefix = (isViewMode && loopTotal > 1) ? `[Vòng ${loopCurrent}/${loopTotal}] ` : '';

                const totalForProg = targetLimit > 0 ? targetLimit : list.length;
                const currentForProg = targetLimit > 0 ? currentSuccess : nextIdx;
                updateProgress(ui, currentForProg, totalForProg, loopPrefix ? `${loopPrefix}${currentForProg}/${totalForProg} cuốn` : '');

                const speedStr = !skipped ? ' (mất ' + fmtTime(elapsedMs / 1000) + ')' : '';
                updateStatus(skipped
                    ? ` ${loopPrefix}Đã bỏ qua cuốn #${index + 1}${targetLimit > 0 ? ' [' + currentSuccess + '/' + targetLimit + ']' : ''} — chuyển sang cuốn tiếp theo`
                    : ` ${loopPrefix}Xong ${targetLimit > 0 ? '[' + currentSuccess + '/' + targetLimit + ' cuốn]' : 'cuốn #' + (index + 1)}${speedStr} — chuyển sang cuốn tiếp theo`, 'success');

                if (!skipped) {
                    await showCountdown(isPerfMode() ? 0 : 3);
                    consumeSkipRequest();
                    if (isStopped()) return;
                }

                window.location.href = list[nextIdx].url;
            };

            setTimeout(runAutoProcess, (isFastMode() || isPerfMode()) ? 120 : 800);
        }

        ui.btnRead.addEventListener('click', async () => {
            if (!(await checkLoginOrAlert(updateStatus))) return;
            setRunningMode(ui);
            try {
                await performAutoReadLogic(updateStatus);
                updateProgress(ui, 1, 1, ' Đã đọc & like xong trang này!');
                clearRetry(window.location.href);
                let history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
                if (!history.includes(window.location.href)) {
                    history.push(window.location.href);
                    localStorage.setItem(KEY_READ_HISTORY, JSON.stringify(history));
                }
            } catch (err) {
                if (err.message === 'TEACHER_BOOK') {
                    updateProgress(ui, 1, 1, ' Sách giáo viên');
                } else if (err.message === 'ALREADY_READ') {
                    updateProgress(ui, 1, 1, ' Cuốn này đã hoàn thành trước đó!');
                } else if (err.message === 'FILE_ACCESS_ERROR') {
                    const handled = handleFileAccessErrorIfNeeded(ui, updateStatus, 'read');
                    if (handled) return;
                } else {
                    updateStatus(_d('IEzhu5dpOiA=') + err.message, 'error');
                }
            }
            setNormalMode(ui);
        });

        ui.btnScan.addEventListener('click', async () => {
            if (!(await checkLoginOrAlert(updateStatus))) return;
            localStorage.removeItem(KEY_SCRAPE_PENDING);
            localStorage.removeItem(KEY_STOP);

            if (!isListPage()) {
                updateStatus(_d('IMSQYW5nIGNodXnhu4NuIHNhbmcgdHJhbmcgZGFuaCBzw6FjaCDEkeG7gyBjw6BvLi4u'), 'info');
                localStorage.setItem(KEY_SCAN_ACTIVE, 'true');
                localStorage.setItem(KEY_SCRAPE_PENDING, 'true');
                window.location.href = LIST_URL;
                return;
            }

            setRunningMode(ui);
            try {
                let raw = await getOrFetchBooks(updateStatus);
                if (!raw || raw.length === 0) {
                    setNormalMode(ui);
                    return;
                }
                localStorage.setItem(KEY_SCAN_LIST, JSON.stringify(raw));
                localStorage.setItem(KEY_SCAN_INDEX, '0');
                localStorage.setItem(KEY_SCAN_RESULT, '[]');
                localStorage.setItem(KEY_SCAN_ACTIVE, 'true');

                updateProgress(ui, 0, raw.length, ' Bắt đầu kiểm tra ' + raw.length + ' sách...');
                await sleep(800);
                window.location.href = raw[0].url;
            } catch (err) {
                setNormalMode(ui);
                localStorage.removeItem(KEY_SCAN_ACTIVE);
                updateStatus(_d('IEzhu5dpIGPDoG86IA==') + err.message, 'error');
            }
        });

        ui.btnAutoAll.addEventListener('click', async () => {
            if (!(await checkLoginOrAlert(updateStatus))) return;
            localStorage.removeItem(KEY_SCRAPE_PENDING);
            localStorage.removeItem(KEY_STOP);
            localStorage.removeItem(KEY_SCAN_ACTIVE);

            let mode = localStorage.getItem(KEY_AUTO_MODE);
            if (!mode) {
                mode = await showAutoModeDialog();
                if (!mode) return;
                localStorage.setItem(KEY_AUTO_MODE, mode);
            }

            if (!isListPage()) {
                updateStatus(_d('IMSQYW5nIGNodXnhu4NuIHNhbmcgZGFuaCBzw6FjaC4uLg=='), 'info');
                localStorage.setItem(KEY_SCRAPE_PENDING, 'true');
                window.location.href = LIST_URL;
                return;
            }

            localStorage.removeItem(KEY_AUTO_MODE);
            localStorage.setItem(KEY_VIEW_COUNT_MODE, mode === '4' ? 'true' : 'false');
            if (mode === '4') {
                updateViewCountDebug({
                    title: 'Đang chuẩn bị danh sách sách',
                    before: null,
                    after: null,
                    observed: null,
                    status: 'Đang chuẩn bị phiên cào lượt đọc...'
                });
            } else {
                updateViewCountDebug();
            }
            localStorage.setItem(KEY_SKIP_COUNT, '0');
            localStorage.setItem(KEY_TEACHER_COUNT, '0');
            localStorage.setItem(KEY_READ_SUCCESS_COUNT, '0');
            localStorage.setItem(KEY_SKIPPED_DETAILS, '[]');

            setRunningMode(ui);
            try {
                let raw = await getOrFetchBooks(updateStatus);
                if (isStopped()) return;
                if (!raw || raw.length === 0) {
                    setNormalMode(ui);
                    return;
                }

                const history = JSON.parse(localStorage.getItem(KEY_READ_HISTORY) || '[]');
                let bookList = [];

                if (mode === '1') {
                    const unreadBooks = raw.filter(b => !history.includes(b.url));
                    if (unreadBooks.length === 0) {
                        updateStatus(_d('IFThuqV0IGPhuqMgc8OhY2ggxJHhu4F1IMSRw6MgxJHhu41jIHLhu5NpISBLaMO0bmcgY8OybiBzw6FjaCBjaMawYSDEkeG7jWMu'), 'success');
                        setNormalMode(ui); return;
                    }
                    const dir = await showDirectionDialog(unreadBooks);
                    if (!dir) { updateStatus(_d('xJDDoyBodeG7tw=='), 'warn'); setNormalMode(ui); return; }
                    bookList = dir.books;
                    localStorage.setItem(KEY_TARGET_COUNT, '0');
                    updateStatus(_d('IFTDrG0gdGjhuqV5IA==') + bookList.length + '/' + raw.length + _d('IGN14buRbiBjaMawYSDEkeG7jWMgKA==') + dir.label + ')', 'info');
                    await sleep(800);
                } else if (mode === '2') {
                    const unreadBooks = raw.filter(b => !history.includes(b.url));
                    if (unreadBooks.length === 0) {
                        updateStatus(_d('IFThuqV0IGPhuqMgc8OhY2ggxJHhu4F1IMSRw6MgxJHhu41jIHLhu5NpISBLaMO0bmcgY8OybiBzw6FjaCBjaMawYSDEkeG7jWMgbsOgby4='), 'success');
                        setNormalMode(ui); return;
                    }
                    const n = await showNumberInputDialog(unreadBooks.length);
                    if (!n) { updateStatus(_d('xJDDoyBodeG7tw=='), 'warn'); setNormalMode(ui); return; }
                    localStorage.setItem(KEY_TARGET_COUNT, n.toString());
                    const dir = await showDirectionDialog(unreadBooks);
                    if (!dir) { updateStatus(_d('xJDDoyBodeG7tw=='), 'warn'); setNormalMode(ui); return; }
                    const ordered = dir.books;
                    bookList = ordered.slice(0, n);
                    const reserve = ordered.slice(n);
                    localStorage.setItem('TR_RESERVE_LIST', JSON.stringify(reserve));
                    updateStatus(_d('IE3hu6VjIHRpw6p1OiDEkOG7jWMgxJHhu6cg') + n + _d('IGN14buRbiBzw6FjaCAo') + dir.label + ')', 'info');
                    await sleep(800);
                } else if (mode === '4') {
                    const loops = await showLoopInputDialog(raw.length);
                    if (!loops) { updateStatus(_d('xJDDoyBodeG7tw=='), 'warn'); setNormalMode(ui); return; }
                    localStorage.setItem(KEY_VIEW_LOOP_TOTAL, loops.toString());
                    localStorage.setItem(KEY_VIEW_LOOP_CURRENT, '1');
                    const dir = await showDirectionDialog(raw);
                    if (!dir) { updateStatus(_d('xJDDoyBodeG7tw=='), 'warn'); setNormalMode(ui); return; }
                    bookList = dir.books;
                    localStorage.setItem(KEY_TARGET_COUNT, '0');
                    updateStatus(_d('IEPDoG8gbMaw4bujdCDEkeG7jWM6IA==') + bookList.length + _d('IGN14buRbiDCtyA=') + loops + _d('IHbDsm5nIGzhurdwICg=') + dir.label + ')', 'info');
                    await sleep(800);
                } else {
                    updateStatus(_d('IE3hu58gYuG6o25nIGNo4buNbiBzw6FjaC4uLg=='), 'info');
                    const pickResult = await showBookPicker(raw, ' Chọn sách cần đọc');
                    if (pickResult === null) { updateStatus(_d('xJDDoyBodeG7tyBjaOG7jW4gc8OhY2g='), 'warn'); setNormalMode(ui); return; }

                    const autoMerged = mergeScanResults();
                    if (autoMerged > 0) { updateStatus(_d('IE7huqFwIA==') + autoMerged + _d('IHPDoWNoIGhvw6BuIHThuqV0'), 'info'); await sleep(600); }

                    const candidate = raw.filter(b => pickResult.selectedUrls.has(b.url));
                    if (candidate.length === 0) { updateStatus(_d('IENoxrBhIGNo4buNbiBzw6FjaCBuw6Bv'), 'warn'); setNormalMode(ui); return; }

                    const dir = await showDirectionDialog(candidate);
                    if (!dir) { updateStatus(_d('xJDDoyBodeG7tw=='), 'warn'); setNormalMode(ui); return; }
                    bookList = dir.books;
                    if (pickResult.limit > 0 && bookList.length > pickResult.limit) bookList = bookList.slice(0, pickResult.limit);
                    localStorage.setItem(KEY_TARGET_COUNT, '0');
                }

                if (bookList.length === 0) { updateStatus(_d('IENoxrBhIGNo4buNbiBzw6FjaCBuw6Bv'), 'warn'); setNormalMode(ui); return; }

                localStorage.setItem(KEY_LIST, JSON.stringify(bookList));
                localStorage.setItem(KEY_INDEX, '0');
                localStorage.setItem(KEY_ACTIVE, 'true');
                localStorage.setItem(KEY_START_TIME, Date.now().toString());

                const targetVal = parseInt(localStorage.getItem(KEY_TARGET_COUNT) || '0', 10);
                const countEst = targetVal > 0 ? targetVal : bookList.length;
                const etaSec = countEst * getAvgSecPerBook();
                updateProgress(ui, 0, countEst, ' Sẵn sàng ' + countEst + ' cuốn — dự kiến ' + fmtTime(etaSec));
                if (mode === '4') {
                    createOrShowViewScrapePopup(ui);
                }
                await sleep(1500);
                if (isStopped()) return;
                window.location.href = bookList[0].url;
            } catch (err) {
                setNormalMode(ui);
                updateStatus(_d('IEzhu5dpOiA=') + err.message, 'error');
            }
        });

        ui.btnStop.addEventListener('click', () => {
            disablePageBlocker();
            closeViewScrapePopup();
            localStorage.removeItem(KEY_SKIP_CURRENT);
            localStorage.setItem(KEY_STOP, 'true');
            const isScanning = localStorage.getItem(KEY_SCAN_ACTIVE) === 'true';
            const isReading = localStorage.getItem(KEY_ACTIVE) === 'true';
            let stopProgressTotal = 0;
            try {
                if (isReading) {
                    const readingBooks = JSON.parse(localStorage.getItem(KEY_LIST) || '[]');
                    const targetCount = parseInt(localStorage.getItem(KEY_TARGET_COUNT) || '0', 10);
                    stopProgressTotal = targetCount > 0 ? targetCount : readingBooks.length;
                } else if (isScanning) {
                    stopProgressTotal = JSON.parse(localStorage.getItem(KEY_SCAN_LIST) || '[]').length;
                }
            } catch (e) { }
            updateProgress(ui, 0, stopProgressTotal);

            if (isReading) {
                localStorage.removeItem(KEY_ACTIVE);
                localStorage.removeItem(KEY_LIST);
                localStorage.removeItem(KEY_INDEX);
                localStorage.removeItem(KEY_LIMIT);
                localStorage.removeItem(KEY_TARGET_COUNT);
                localStorage.removeItem(KEY_VIEW_COUNT_MODE);
                localStorage.removeItem(KEY_VIEW_LOOP_TOTAL);
                localStorage.removeItem(KEY_VIEW_LOOP_CURRENT);
                localStorage.removeItem(KEY_START_TIME);
                const overlay = document.getElementById('tr-countdown-overlay');
                if (overlay) overlay.remove();
                updateStatus(_d('IMSQw6MgZOG7q25nIMSR4buNYyBzw6FjaA=='), 'warn');
                ui.etaWrap.classList.remove('tr-on');
                setNormalMode(ui);
                setTimeout(() => localStorage.removeItem(KEY_STOP), 1000);
            } else if (isScanning) {
                localStorage.removeItem(KEY_SCAN_ACTIVE);
                localStorage.removeItem(KEY_SCAN_LIST);
                localStorage.removeItem(KEY_SCAN_INDEX);
                updateStatus(_d('IMSQw6MgZOG7q25nIHF1w6l0'), 'warn');
                setNormalMode(ui);
                setTimeout(() => localStorage.removeItem(KEY_STOP), 1000);
            } else {
                updateStatus(_d('IMSQYW5nIGThu6tuZyBjw6BvLi4u'), 'warn');
            }
        });

        ui.btnSkip.addEventListener('click', () => {
            if (localStorage.getItem(KEY_ACTIVE) !== 'true' || ui.btnSkip.disabled) return;
            ui.btnSkip.disabled = true;
            localStorage.setItem(KEY_SKIP_CURRENT, 'true');
            updateStatus(_d('IMSQYW5nIGLhu48gcXVhIGN14buRbiBoaeG7h24gdOG6oWkuLi4='), 'warn');
        });
    };

    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        setTimeout(initApp, 100);
    } else {
        window.addEventListener('DOMContentLoaded', initApp);
    }
})();
