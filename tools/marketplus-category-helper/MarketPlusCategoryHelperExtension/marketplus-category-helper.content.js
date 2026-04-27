// ==UserScript==
// @name         MarketPlus Category Helper
// @namespace    cafe24-marketplus-helper
// @version      4.4
// @description  마켓플러스 카테고리 네이버연동 자동매칭
// @match        *://mp.cafe24.com/*registerall*
// @match        *://*.cafe24.com/mp/product/front/registerall*
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
    'use strict';

    var STORAGE_KEY = 'mp_category_presets';
    var PROXY_URL = 'http://localhost:5555';
    var naverContext = null; // 네이버 카테고리 맥락 저장
    var categoryMapAutoApplied = false;

    function sleep(ms) {
        return new Promise(function (resolve) { setTimeout(resolve, ms); });
    }

    // ─── 네이버 카테고리 API ───
    async function fetchNaverCategory(productName) {
        try {
            var resp = await fetch(PROXY_URL + '/?q=' + encodeURIComponent(productName));
            return await resp.json();
        } catch (e) {
            return { category: '', levels: [], error: 'proxy not running' };
        }
    }

    // ─── 상품명 ───
    function getProductNames() {
        var names = [];
        var inputs = document.querySelectorAll('input[type="hidden"]');
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].name && inputs[i].name.indexOf('product_name') >= 0 && inputs[i].value) {
                names.push(inputs[i].value);
            }
        }
        return names;
    }

    function getProductInfos() {
        var products = [];
        var inputs = document.querySelectorAll('input[type="hidden"]');
        for (var i = 0; i < inputs.length; i++) {
            var input = inputs[i];
            if (!input.name || input.name.indexOf('product_name') < 0 || !input.value) continue;
            var code = '';
            var m = input.name.match(/prd_info\[([^\]]+)\]\[product_name\]/);
            if (m) code = m[1];
            products.push({ code: code, name: input.value });
        }
        if (products.length === 0) {
            var names = getProductNames();
            for (var n = 0; n < names.length; n++) products.push({ code: '', name: names[n] });
        }
        return products;
    }

    async function fetchCategoryMap(product) {
        try {
            var url = PROXY_URL + '/api/category-map?productName=' + encodeURIComponent(product.name || '') +
                '&productCode=' + encodeURIComponent(product.code || '');
            var resp = await fetch(url);
            return await resp.json();
        } catch (e) {
            return { matched: false, error: 'category map server not running' };
        }
    }

    function readFileAsBase64(file) {
        return new Promise(function (resolve, reject) {
            var reader = new FileReader();
            reader.onload = function () {
                var text = String(reader.result || '');
                resolve(text.indexOf(',') >= 0 ? text.split(',')[1] : text);
            };
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    async function uploadCategoryMapFile(file) {
        if (!file) {
            setStatus('map', '업로드할 xlsx 파일을 선택하세요');
            return;
        }
        setStatus('map', '카테고리맵 업로드 중...');
        try {
            var contentBase64 = await readFileAsBase64(file);
            var resp = await fetch(PROXY_URL + '/api/upload-map', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filename: file.name, contentBase64: contentBase64 })
            });
            var data = await resp.json();
            if (!data.ok) {
                setStatus('map', '업로드 실패: ' + (data.error || 'unknown error'));
                return;
            }
            categoryMapAutoApplied = false;
            setStatus('map', '업로드 완료: 상품 ' + data.productCount + '개 / 레코드 ' + data.recordCount + '개');
            setTimeout(function () { applyCategoryMap(false); }, 400);
        } catch (e) {
            setStatus('map', '업로드 실패: ' + e.message);
        }
    }

    // ─── 키워드 생성 ───
    function generateKeywords(productName) {
        var words = productName.split(/\s+/);
        var meaningful = [];

        for (var i = 0; i < words.length; i++) {
            var w = words[i];
            if (w.length < 2) continue;
            if (/^[0-9A-Za-z]+$/.test(w) && w.length <= 3) continue;
            if (/^(고정|방지|이중|합금강|합금|대용량|멀티|미니|휴대용|고급|특가|무료배송|스프링식|낙하|조절|편하는|빠른|부위별|맞춤)$/.test(w)) continue;
            meaningful.push(w);
        }

        var keywords = [];
        var seen = {};
        function add(k) { if (!k || seen[k]) return; seen[k] = true; keywords.push(k); }

        if (meaningful.length >= 2) add(meaningful[0] + ' ' + meaningful[1]);
        for (var j = 0; j < Math.min(meaningful.length, 5); j++) add(meaningful[j]);
        return keywords;
    }

    // ─── 마켓 상태 ───
    function getMarketContainers() {
        var c = document.querySelectorAll('td.eTdCategoryTemplate[data-marketkey]');
        if (c.length === 0) c = document.querySelectorAll('td[data-marketkey]');
        return c;
    }

    function marketIdKey(k) { return k.replace(/\|/g, '_'); }

    function isMarketMatched(mk) {
        var sel = document.getElementById('eMarketCategory1_' + marketIdKey(mk));
        return sel && sel.value && sel.value !== '';
    }

    function getUnmatchedCount() {
        var n = 0;
        getMarketContainers().forEach(function (td) {
            var mk = td.getAttribute('data-marketkey');
            if (mk && !isMarketMatched(mk)) n++;
        });
        return n;
    }

    function getTotalMarketCount() { return getMarketContainers().length; }

    // ─── 맥락 기반 카테고리 스코어링 ───
    function scoreCategoryMatch(categoryText, productName) {
        var words = productName.split(/\s+/);
        var score = 0;
        var catText = categoryText.replace(/\s/g, '');

        for (var i = 0; i < words.length; i++) {
            var w = words[i];
            if (w.length >= 2 && catText.indexOf(w) >= 0) {
                score++;
                if (i < 3) score += 0.5;
            }
        }

        // 네이버 맥락이 있으면 카테고리 도메인 체크
        if (naverContext && naverContext.topLevel) {
            var topWords = naverContext.topLevel.split('/');
            var hasContext = false;
            for (var t = 0; t < topWords.length; t++) {
                if (topWords[t].length >= 2 && catText.indexOf(topWords[t]) >= 0) {
                    hasContext = true;
                    score += 3;
                }
            }

            // 네이버 카테고리의 중간 레벨과 매칭되면 추가 보너스
            if (naverContext.levels) {
                for (var lv = 1; lv < naverContext.levels.length; lv++) {
                    var lvWord = naverContext.levels[lv];
                    if (lvWord && lvWord.length >= 2 && catText.indexOf(lvWord) >= 0) {
                        score += 2;
                    }
                }
            }

            // 맥락과 완전히 다른 도메인이면 큰 감점
            if (!hasContext && score > 0) {
                var wrongDomains = [
                    '골프', '식품', '냉장', '냉동', '키보드', '마우스', '프린터',
                    '오일', '자동차', '타이어', '골프클럽', '해외직구', '중고',
                    '반려', '애완', '낚시', '캠핑'
                ];
                for (var wd = 0; wd < wrongDomains.length; wd++) {
                    if (catText.indexOf(wrongDomains[wd]) >= 0) {
                        score -= 5;
                    }
                }
            }
        }

        return score;
    }

    // ─── 미매칭 마켓 자동 클릭 ───
    function autoMatchUnmatched(productName) {
        var matched = 0;
        var results = [];
        getMarketContainers().forEach(function (td) {
            var mk = td.getAttribute('data-marketkey');
            if (!mk || isMarketMatched(mk)) return;

            var links = td.querySelectorAll('a.txtLink');
            if (links.length === 0) return;

            var bestLink = null;
            var bestScore = -999;

            for (var i = 0; i < links.length; i++) {
                var score = scoreCategoryMatch(links[i].textContent.trim(), productName);
                if (score > bestScore) {
                    bestScore = score;
                    bestLink = links[i];
                }
            }

            if (bestLink && bestScore > 0) {
                bestLink.click();
                matched++;
                results.push(mk.split('|')[0] + ': ' + bestLink.textContent.trim());
            }
        });
        return { count: matched, details: results };
    }

    // ─── 표준카테고리 검색 ───
    function findSearchInput() {
        var inputs = document.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].placeholder && inputs[i].placeholder.indexOf('카테고리') >= 0)
                return inputs[i];
        }
        return null;
    }

    function triggerSearchAction(keyword) {
        var input = findSearchInput();
        if (!input) return false;
        var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, keyword);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.focus();
        setTimeout(function () {
            input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
            input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
        }, 200);
        return true;
    }

    // ─── 자동 매칭 (네이버 카테고리 계층 → 다중 키워드) ───
    async function runAutoMatch() {
        var names = getProductNames();
        if (names.length === 0) { setStatus('match', '상품명을 찾을 수 없습니다'); return; }

        var productName = names[0];
        var totalMarkets = getTotalMarketCount();
        var matchBtn = document.getElementById('mph-auto-match');
        matchBtn.disabled = true;
        matchBtn.textContent = '매칭 중...';

        // 1. 네이버에서 카테고리 계층 가져오기
        setStatus('match', '네이버에서 카테고리 검색 중...');
        var naver = await fetchNaverCategory(productName);

        var keywords = [];

        if (naver.category && naver.levels && naver.levels.length > 0) {
            naverContext = naver; // 맥락 저장

            var naverInfo = document.getElementById('mph-naver-info');
            if (naverInfo) {
                naverInfo.textContent = naver.fullPath || naver.category;
                naverInfo.style.display = 'block';
            }

            // 네이버 카테고리 계층을 역순(구체→일반)으로 키워드 생성
            // 예: ['화장품/미용','스킨케어','페이스소품','리프팅밴드']
            //   → 검색순서: 리프팅밴드 → 페이스소품 → 스킨케어 → 화장품
            for (var lv = naver.levels.length - 1; lv >= 0; lv--) {
                var lvKeyword = naver.levels[lv];
                if (lvKeyword && lvKeyword.length >= 2) {
                    // "화장품/미용" → "화장품" 과 "미용" 분리
                    var subWords = lvKeyword.split('/');
                    for (var sw = 0; sw < subWords.length; sw++) {
                        if (subWords[sw].length >= 2 && keywords.indexOf(subWords[sw]) < 0) {
                            keywords.push(subWords[sw]);
                        }
                    }
                }
            }

            setStatus('match', '네이버: ' + naver.fullPath);
            document.getElementById('mph-keyword').value = keywords[0] || '';
            await sleep(500);
        } else {
            naverContext = null;
            if (naver.error === 'proxy not running') {
                setStatus('match', '네이버 프록시 미실행 - 상품명 키워드로 진행');
            } else {
                setStatus('match', '네이버 결과 없음 - 상품명 키워드로 진행');
            }
            await sleep(300);
        }

        // 상품명 키워드를 뒤에 추가
        var productKeywords = generateKeywords(productName);
        for (var pk = 0; pk < productKeywords.length; pk++) {
            if (keywords.indexOf(productKeywords[pk]) < 0) {
                keywords.push(productKeywords[pk]);
            }
        }

        // 2. 순차 검색 + 매칭
        var allResults = [];

        for (var i = 0; i < keywords.length; i++) {
            var unmatched = getUnmatchedCount();
            if (unmatched === 0) break;

            var kw = keywords[i];
            setStatus('match', '(' + (i + 1) + '/' + keywords.length + ') "' + kw + '" 검색 중... (남은: ' + unmatched + ')');

            triggerSearchAction(kw);
            await sleep(2500);

            var result = autoMatchUnmatched(productName);
            if (result.count > 0) {
                allResults = allResults.concat(result.details);
                var done = totalMarkets - getUnmatchedCount();
                setStatus('match', '"' + kw + '" → ' + result.count + '개 (' + done + '/' + totalMarkets + ')');
            }
            await sleep(800);
        }

        // 3. 결과
        matchBtn.disabled = false;
        matchBtn.textContent = '자동 매칭 실행';

        var finalDone = totalMarkets - getUnmatchedCount();
        var remaining = getUnmatchedCount();

        var msg = finalDone + '/' + totalMarkets + ' 마켓 매칭!';
        if (remaining > 0) msg += ' (미매칭 ' + remaining + '개는 수동)';
        else msg += ' 전체 완료!';
        setStatus('match', msg);

        var resultEl = document.getElementById('mph-match-result');
        if (resultEl && allResults.length > 0) {
            resultEl.style.display = 'block';
            resultEl.textContent = allResults.join('\n');
        }

        setTimeout(refreshSummary, 1000);
    }

    // ─── 프리셋 ───
    function getMarketKeys() {
        var selects = document.querySelectorAll('select[name^="market_category1["]');
        var keys = [];
        selects.forEach(function (sel) {
            var m = sel.name.match(/market_category1\[(.+?)\]/);
            if (m) keys.push(m[1]);
        });
        return keys;
    }

    function readCurrentState() {
        var state = {};
        getMarketKeys().forEach(function (key) {
            var idKey = marketIdKey(key);
            var depths = {};
            for (var d = 1; d <= 7; d++) {
                var sel = document.getElementById('eMarketCategory' + d + '_' + idKey);
                if (sel && sel.value) {
                    depths[d] = { value: sel.value, text: (sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : '') };
                }
            }
            if (Object.keys(depths).length > 0) state[key] = depths;
        });
        return state;
    }

    function waitForOption(sel, val, timeout) {
        return new Promise(function (resolve) {
            var start = Date.now();
            (function check() {
                for (var i = 0; i < sel.options.length; i++) {
                    if (String(sel.options[i].value) === String(val)) { resolve(true); return; }
                }
                if (Date.now() - start > timeout) { resolve(false); return; }
                setTimeout(check, 150);
            })();
        });
    }

    async function applyPresetForMarket(mk, depths) {
        var idKey = marketIdKey(mk);
        var keys = Object.keys(depths).map(Number).sort(function (a, b) { return a - b; });
        var maxD = keys[keys.length - 1];
        for (var d = 1; d <= maxD; d++) {
            var dd = depths[d]; if (!dd) continue;
            var sel = document.getElementById('eMarketCategory' + d + '_' + idKey);
            if (!sel) break;
            var found = await waitForOption(sel, dd.value, 5000);
            if (!found) break;
            sel.value = dd.value;
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            await sleep(500);
        }
    }

    async function applyPreset(state) {
        var ps = [];
        for (var mk in state) { if (state.hasOwnProperty(mk)) ps.push(applyPresetForMarket(mk, state[mk])); }
        await Promise.all(ps);
    }

    function normalizeMarketName(raw) {
        var s = String(raw || '').toLowerCase();
        if (s.indexOf('auction') >= 0 || s.indexOf('옥션') >= 0) return 'auction';
        if (s.indexOf('coupang') >= 0 || s.indexOf('쿠팡') >= 0) return 'coupang';
        if (s.indexOf('gmarket') >= 0 || s.indexOf('g마켓') >= 0 || s.indexOf('g마') >= 0) return 'gmarket';
        if (s.indexOf('11st') >= 0 || s.indexOf('11번') >= 0) return '11st';
        if (s.indexOf('smart') >= 0 || s.indexOf('naver') >= 0 || s.indexOf('ncp') >= 0 || s.indexOf('스마트') >= 0) return 'naver';
        if (s.indexOf('lotte') >= 0 || s.indexOf('롯데') >= 0) return 'lotteon';
        if (s.indexOf('kakao') >= 0 || s.indexOf('카카오') >= 0) return 'kakao';
        return s.split('|')[0];
    }

    function getMarketAliasForContainer(td) {
        var mk = td.getAttribute('data-marketkey') || '';
        var direct = normalizeMarketName(mk.split('|')[0]);
        if (direct) return direct;
        return normalizeMarketName(td.textContent || '');
    }

    function getMappedMarket(markets, alias) {
        if (!markets) return null;
        if (markets[alias]) return markets[alias];
        if (alias === 'naver' && markets.smartstore) return markets.smartstore;
        if (alias === 'smartstore' && markets.naver) return markets.naver;
        if (alias === '11st' && markets['11번가']) return markets['11번가'];
        return null;
    }

    function parseSelectorJson(value) {
        if (!value) return null;
        if (typeof value === 'object') return value;
        try { return JSON.parse(value); } catch (e) { return null; }
    }

    function getCategorySegments(path) {
        return String(path || '').split(/[>\/]+/).map(function (x) { return x.trim(); }).filter(Boolean);
    }

    function getCategorySearchKeyword(path) {
        var segments = getCategorySegments(path);
        if (segments.length === 0) return '';
        return segments[segments.length - 1];
    }

    function normalizeCategoryText(value) {
        return String(value || '').toLowerCase().replace(/\s+/g, '').replace(/[>\/|,-]/g, '');
    }

    function scoreMappedCategoryLink(linkText, categoryPath) {
        var cat = normalizeCategoryText(linkText);
        var segments = getCategorySegments(categoryPath);
        var score = 0;
        if (!cat || segments.length === 0) return score;

        var leaf = normalizeCategoryText(segments[segments.length - 1]);
        if (leaf && cat.indexOf(leaf) >= 0) score += 8;

        for (var i = 0; i < segments.length; i++) {
            var seg = normalizeCategoryText(segments[i]);
            if (seg && cat.indexOf(seg) >= 0) score += i === segments.length - 1 ? 3 : 2;
        }
        return score;
    }

    function clickBestMappedLink(td, categoryPath) {
        var links = td.querySelectorAll('a.txtLink');
        var bestLink = null;
        var bestScore = 0;
        for (var i = 0; i < links.length; i++) {
            var score = scoreMappedCategoryLink(links[i].textContent || '', categoryPath);
            if (score > bestScore) {
                bestScore = score;
                bestLink = links[i];
            }
        }
        if (bestLink && bestScore >= 8) {
            bestLink.click();
            return bestLink.textContent.trim();
        }
        return '';
    }

    async function applyCategoryMap(auto) {
        if (auto && categoryMapAutoApplied) return;

        var products = getProductInfos();
        if (products.length === 0) {
            setStatus('map', '상품명을 찾을 수 없습니다');
            return;
        }

        var product = products[0];
        setStatus('map', '매칭표 조회 중...');
        var result = await fetchCategoryMap(product);

        if (!result.matched) {
            setStatus('map', '매칭표 없음: ' + (result.error || '상품명 미일치'));
            return;
        }

        var applied = 0;
        var skipped = 0;
        var pendingSearch = [];
        var containers = getMarketContainers();
        for (var i = 0; i < containers.length; i++) {
            var td = containers[i];
            var mk = td.getAttribute('data-marketkey');
            if (!mk || isMarketMatched(mk)) { skipped++; continue; }

            var alias = getMarketAliasForContainer(td);
            var info = getMappedMarket(result.markets, alias);
            var depths = info ? parseSelectorJson(info.selectorJson) : null;
            if (!info) { skipped++; continue; }

            if (depths) {
                await applyPresetForMarket(mk, depths);
                if (isMarketMatched(mk)) applied++;
                else skipped++;
            } else if (info.categoryPath) {
                pendingSearch.push({ td: td, mk: mk, info: info, keyword: getCategorySearchKeyword(info.categoryPath) });
            } else {
                skipped++;
            }
        }

        var searched = {};
        for (var p = 0; p < pendingSearch.length; p++) {
            var item = pendingSearch[p];
            if (!item.keyword) { skipped++; continue; }
            if (!searched[item.keyword]) {
                setStatus('map', '매칭표 검색 중: ' + item.keyword);
                triggerSearchAction(item.keyword);
                await sleep(2500);
                searched[item.keyword] = true;
            }
            if (isMarketMatched(item.mk)) { skipped++; continue; }
            var clicked = clickBestMappedLink(item.td, item.info.categoryPath);
            if (clicked) applied++;
            else skipped++;
        }

        categoryMapAutoApplied = true;
        var note = '';
        if (result.reviewNeeded || Number(result.confidence || 0) < 0.8) note = ' / 검수필요';
        setStatus('map', '매칭표 적용: ' + applied + '개, 건너뜀 ' + skipped + '개' + note + ' (' + result.productName + ')');
        setTimeout(refreshSummary, 800);
    }

    function loadPresets() { try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); } catch (e) { return {}; } }
    function savePresets(p) { localStorage.setItem(STORAGE_KEY, JSON.stringify(p)); }
    function presetSummary(state) {
        var parts = [];
        for (var k in state) { if (!state.hasOwnProperty(k)) continue; var d = state[k]; var ks = Object.keys(d).map(Number); var m = Math.max.apply(null, ks); if (d[m]) parts.push(d[m].text); }
        return parts.slice(0, 2).join(', ');
    }

    // ─── UI ───
    function setStatus(s, msg) { var el = document.getElementById('mph-status-' + s); if (el) el.textContent = msg; }

    function refreshSummary() {
        var el = document.getElementById('mph-current-summary'); if (!el) return;
        var state = readCurrentState();
        if (Object.keys(state).length === 0) { el.textContent = '카테고리 미선택'; return; }
        var lines = [];
        for (var k in state) {
            if (!state.hasOwnProperty(k)) continue;
            var path = Object.keys(state[k]).sort(function (a, b) { return a - b; }).map(function (d) { return state[k][d].text; }).join(' > ');
            lines.push(k.split('|')[0] + ': ' + path);
        }
        el.textContent = lines.join('\n');
    }

    function refreshPresetSelect() {
        var sel = document.getElementById('mph-preset-select'); if (!sel) return;
        var presets = loadPresets();
        sel.innerHTML = '<option value="">-- 선택 --</option>';
        for (var name in presets) {
            if (!presets.hasOwnProperty(name)) continue;
            var opt = document.createElement('option'); opt.value = name;
            opt.textContent = name + ' (' + presetSummary(presets[name]) + ')';
            sel.appendChild(opt);
        }
    }

    function createUI() {
        if (document.getElementById('mph-panel')) return;
        var productNames = getProductNames();
        var productDisplay = productNames.length > 0 ? productNames[0].substring(0, 40) + (productNames[0].length > 40 ? '...' : '') : '(없음)';
        var defaultKw = productNames.length > 0 ? generateKeywords(productNames[0])[0] || '' : '';

        var panel = document.createElement('div'); panel.id = 'mph-panel';
        var header = document.createElement('div'); header.id = 'mph-header';
        header.innerHTML = '<span id="mph-title">Category Helper v4</span><span id="mph-toggle">▼</span>';
        panel.appendChild(header);

        var body = document.createElement('div'); body.id = 'mph-body';
        body.innerHTML = [
            '<div class="mph-section">',
            '  <div class="mph-label">상품명</div>',
            '  <div class="mph-product-name">' + productDisplay + '</div>',
            '  <button id="mph-naver-btn" class="mph-btn mph-btn-naver" style="width:100%;margin-top:4px;">네이버쇼핑에서 확인</button>',
            '</div>',
            '<div class="mph-section">',
            '  <div class="mph-label">자동 매칭</div>',
            '  <div id="mph-naver-info" class="mph-naver-info" style="display:none;"></div>',
            '  <div class="mph-row">',
            '    <input type="text" id="mph-keyword" placeholder="키워드" value="' + defaultKw + '">',
            '    <button id="mph-search-btn" class="mph-btn mph-btn-blue">검색</button>',
            '  </div>',
            '  <div class="mph-hint">네이버 카테고리 계층(구체→일반) + 상품명 키워드로 순차 검색</div>',
            '  <button id="mph-auto-match" class="mph-btn mph-btn-auto" style="width:100%;padding:10px;font-size:14px;margin-top:6px;">자동 매칭 실행</button>',
            '  <div class="mph-status" id="mph-status-match"></div>',
            '  <div id="mph-match-result" class="mph-summary" style="display:none;margin-top:6px;"></div>',
            '</div>',
            '<div class="mph-section">',
            '  <div class="mph-label">카테고리맵</div>',
            '  <input id="mph-map-file" type="file" accept=".xlsx" style="display:none;">',
            '  <div class="mph-row">',
            '    <button id="mph-map-file-pick" class="mph-btn mph-btn-blue" style="flex:1;">파일 선택</button>',
            '    <button id="mph-map-upload" class="mph-btn mph-btn-gray">업로드</button>',
            '  </div>',
            '  <div id="mph-map-file-name" class="mph-hint">선택된 파일 없음</div>',
            '  <button id="mph-map-apply" class="mph-btn mph-btn-green" style="width:100%;padding:9px;font-size:13px;">업로드 매칭표 적용</button>',
            '  <div class="mph-hint">localhost:5555에 업로드된 엑셀 매칭표를 상품명으로 조회</div>',
            '  <div class="mph-status" id="mph-status-map"></div>',
            '</div>',
            '<div class="mph-section">',
            '  <div class="mph-label">프리셋</div>',
            '  <div class="mph-row"><select id="mph-preset-select"><option value="">-- 선택 --</option></select></div>',
            '  <div class="mph-row">',
            '    <button id="mph-preset-apply" class="mph-btn mph-btn-green">적용</button>',
            '    <button id="mph-preset-delete" class="mph-btn mph-btn-red">삭제</button>',
            '  </div>',
            '  <div class="mph-divider"></div>',
            '  <div class="mph-row">',
            '    <input type="text" id="mph-preset-name" placeholder="프리셋 이름">',
            '    <button id="mph-preset-save" class="mph-btn mph-btn-blue">저장</button>',
            '  </div>',
            '  <div class="mph-status" id="mph-status-preset"></div>',
            '</div>',
            '<div class="mph-section mph-section-last">',
            '  <div class="mph-label">현재 카테고리</div>',
            '  <div id="mph-current-summary" class="mph-summary">카테고리 미선택</div>',
            '  <button id="mph-refresh-summary" class="mph-btn mph-btn-gray">새로고침</button>',
            '</div>'
        ].join('');
        panel.appendChild(body);

        if (!document.getElementById('mph-style')) {
            var style = document.createElement('style');
            style.id = 'mph-style';
            style.textContent = '#mph-panel{position:fixed;top:12px;right:12px;z-index:999999;width:320px;background:#fff;border:1px solid #d0d7de;border-radius:10px;box-shadow:0 6px 20px rgba(0,0,0,.12);font-family:"Malgun Gothic",sans-serif;font-size:13px;color:#24292f;overflow:hidden}#mph-header{display:flex;justify-content:space-between;align-items:center;padding:10px 14px;background:#2c6fbb;color:#fff;cursor:move;user-select:none}#mph-title{font-weight:700;font-size:13px}#mph-toggle{cursor:pointer;font-size:12px}#mph-body{padding:12px 14px;max-height:80vh;overflow-y:auto}.mph-section{margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid #eaeef2}.mph-section-last{border-bottom:none;margin-bottom:0;padding-bottom:0}.mph-label{font-weight:700;margin-bottom:6px;color:#1f2328;font-size:12px}.mph-product-name{font-size:11px;color:#57606a;padding:4px 0;word-break:break-all}.mph-hint{font-size:10px;color:#8b949e;margin-top:2px}.mph-naver-info{font-size:11px;color:#03c75a;font-weight:700;padding:4px 8px;background:#e8f8ee;border-radius:4px;margin-bottom:6px}.mph-row{display:flex;gap:6px;margin-bottom:6px;align-items:center}.mph-row input[type="text"],.mph-row select{flex:1;padding:5px 8px;border:1px solid #d0d7de;border-radius:5px;font-size:12px;outline:none}.mph-btn{padding:5px 10px;border:none;border-radius:5px;cursor:pointer;font-size:11px;font-weight:600;white-space:nowrap}.mph-btn:hover{opacity:.85}.mph-btn:disabled{opacity:.5;cursor:not-allowed}.mph-btn-blue{background:#2c6fbb;color:#fff}.mph-btn-green{background:#2da44e;color:#fff}.mph-btn-red{background:#cf222e;color:#fff}.mph-btn-gray{background:#6e7781;color:#fff}.mph-btn-naver{background:#03c75a;color:#fff;font-size:12px;font-weight:700}.mph-btn-auto{background:#8250df;color:#fff;font-weight:700;border-radius:6px}.mph-divider{height:1px;background:#eaeef2;margin:8px 0}.mph-status{font-size:11px;color:#57606a;margin-top:4px;min-height:16px}.mph-summary{font-size:11px;color:#57606a;background:#f6f8fa;border-radius:4px;padding:6px 8px;margin-bottom:6px;max-height:150px;overflow-y:auto;line-height:1.5;white-space:pre-wrap}';
            document.head.appendChild(style);
        }
        document.body.appendChild(panel);

        // drag
        var isDragging = false, dx = 0, dy = 0;
        header.addEventListener('mousedown', function (e) { if (e.target.id === 'mph-toggle') return; isDragging = true; dx = e.clientX - panel.getBoundingClientRect().left; dy = e.clientY - panel.getBoundingClientRect().top; });
        document.addEventListener('mousemove', function (e) { if (!isDragging) return; panel.style.left = (e.clientX - dx) + 'px'; panel.style.top = (e.clientY - dy) + 'px'; panel.style.right = 'auto'; });
        document.addEventListener('mouseup', function () { isDragging = false; });

        document.getElementById('mph-toggle').addEventListener('click', function () {
            var b = document.getElementById('mph-body');
            if (b.style.display === 'none') { b.style.display = ''; this.textContent = '▼'; } else { b.style.display = 'none'; this.textContent = '▶'; }
        });

        document.getElementById('mph-naver-btn').addEventListener('click', function () {
            var n = getProductNames(); if (n.length > 0) window.open('https://search.shopping.naver.com/search/all?query=' + encodeURIComponent(n[0]), '_blank');
        });

        document.getElementById('mph-search-btn').addEventListener('click', function () {
            var kw = document.getElementById('mph-keyword').value.trim();
            if (kw) { triggerSearchAction(kw); setStatus('match', '"' + kw + '" 검색됨'); }
        });
        document.getElementById('mph-keyword').addEventListener('keydown', function (e) { if (e.key === 'Enter') document.getElementById('mph-search-btn').click(); });

        document.getElementById('mph-auto-match').addEventListener('click', function () { runAutoMatch(); });
        document.getElementById('mph-map-file-pick').addEventListener('click', function () { document.getElementById('mph-map-file').click(); });
        document.getElementById('mph-map-file').addEventListener('change', function () {
            var file = this.files && this.files[0];
            document.getElementById('mph-map-file-name').textContent = file ? file.name : '선택된 파일 없음';
        });
        document.getElementById('mph-map-upload').addEventListener('click', function () {
            var input = document.getElementById('mph-map-file');
            uploadCategoryMapFile(input.files && input.files[0]);
        });
        document.getElementById('mph-map-apply').addEventListener('click', function () { applyCategoryMap(false); });

        refreshPresetSelect();
        document.getElementById('mph-preset-save').addEventListener('click', function () {
            var name = document.getElementById('mph-preset-name').value.trim();
            if (!name) { setStatus('preset', '이름을 입력하세요'); return; }
            var state = readCurrentState();
            if (Object.keys(state).length === 0) { setStatus('preset', '선택된 카테고리 없음'); return; }
            var presets = loadPresets(); presets[name] = state; savePresets(presets);
            refreshPresetSelect(); document.getElementById('mph-preset-name').value = '';
            setStatus('preset', '"' + name + '" 저장 완료');
        });
        document.getElementById('mph-preset-apply').addEventListener('click', async function () {
            var name = document.getElementById('mph-preset-select').value;
            if (!name) { setStatus('preset', '선택하세요'); return; }
            var presets = loadPresets(); if (!presets[name]) return;
            setStatus('preset', '적용 중...'); await applyPreset(presets[name]);
            setStatus('preset', '적용 완료'); refreshSummary();
        });
        document.getElementById('mph-preset-delete').addEventListener('click', function () {
            var name = document.getElementById('mph-preset-select').value; if (!name) return;
            if (!confirm('"' + name + '" 삭제?')) return;
            var presets = loadPresets(); delete presets[name]; savePresets(presets);
            refreshPresetSelect(); setStatus('preset', '삭제 완료');
        });

        document.getElementById('mph-refresh-summary').addEventListener('click', refreshSummary);
        document.addEventListener('change', function (e) { if (e.target && e.target.classList && e.target.classList.contains('category')) setTimeout(refreshSummary, 600); });
        setTimeout(refreshSummary, 1000);
        setTimeout(function () { applyCategoryMap(true); }, 1500);
    }

    function ensureUI() {
        if (!document.body) return;
        try {
            createUI();
        } catch (e) {
            console.error('[MarketPlus Category Helper] UI 생성 실패', e);
        }
    }

    function boot() {
        ensureUI();

        var tries = 0;
        var timer = setInterval(function () {
            tries++;
            ensureUI();
            if (document.getElementById('mph-panel') && tries >= 10) clearInterval(timer);
            if (tries >= 120) clearInterval(timer);
        }, 500);

        if (window.MutationObserver && document.documentElement) {
            var observer = new MutationObserver(function () {
                if (!document.getElementById('mph-panel')) ensureUI();
            });
            observer.observe(document.documentElement, { childList: true, subtree: true });
        }
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
    else boot();
})();
