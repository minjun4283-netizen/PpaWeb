(function () {
  try {
    if (window.__PPA_FORM_LOADED__) return;
    window.__PPA_FORM_LOADED__ = true;

    var SCHEMA = null;
    var SCHEMA_BY_KEY = {};

    var TABLE_META = {
      "T_발전소": { label: "발전소" },
      "T_구매계약": { label: "구매계약" },
      "T_수요기업": { label: "수요기업" },
      "T_판매계약": { label: "판매계약" },
      "T_전기사용지": { label: "전기사용지" },
      "T_수급매칭": { label: "수급매칭" }
    };

    var FIELD_LOOKUP = {
      "T_구매계약": { "발전소ID": "T_발전소" },
      "T_판매계약": { "수요기업ID": "T_수요기업" },
      "T_전기사용지": { "판매계약ID": "T_판매계약" },
      "T_수급매칭": { "전기사용지ID": "T_전기사용지", "구매계약ID": "T_구매계약" }
    };

    var optionCache = {};
    var recordCache = {};

    // 저장/삭제할 때 "누가 했는지"를 이력에 남기기 위한 표시 이름 - 한 번
    // 정해두면 이 브라우저에서는 계속 기억합니다. 안 정하면 서버가 자동으로
    // 이 컴퓨터의 Windows 로그인 계정으로 남깁니다(그래도 무방하지만, 로그인
    // 계정이 사람 이름이 아닐 수 있어 원하면 여기서 바꿀 수 있게 함).
    var ACTOR_STORAGE_KEY = "ppa_actor_name";
    function getActorName() {
      try { return (localStorage.getItem(ACTOR_STORAGE_KEY) || "").trim(); } catch (e) { return ""; }
    }
    function setActorName(v) {
      try { localStorage.setItem(ACTOR_STORAGE_KEY, (v || "").trim()); } catch (e) { /* 무시 */ }
    }

    // loadedPk: 지금 폼이 "기존 데이터를 불러온" 상태인지(삭제 가능) 아니면
    // "새 입력" 상태인지(삭제 버튼 비활성) 추적합니다.
    var formState = { loadedPk: null };

    // appMode: "single"(개별입력) | "groupA" | "groupB" - 지금 어느 화면을
    // 보여주고 있는지. 그룹 모드에서는 mainSel/fields 대신 groupWrap을 씁니다.
    var appMode = "single";
    var groupState = null;

    // 저장하지 않은 편집 내용이 있는지 - 있으면 화면 전환/닫기 전에 한 번
    // 물어봐서 실수로 입력한 내용이 사라지는 일을 막습니다.
    var formDirty = false;
    var groupDirty = false;
    function hasUnsavedChanges() {
      return appMode === "single" ? formDirty : groupDirty;
    }
    function confirmDiscardIfDirty(message) {
      if (!hasUnsavedChanges()) return true;
      return window.confirm(message || "저장하지 않은 변경 내용이 있습니다. 계속하면 사라집니다. 계속할까요?");
    }

    var GROUP_DEFS = {
      groupA: {
        key: "groupA",
        label: "그룹A: 발전소 + 구매계약",
        master: "T_발전소",
        child: { table: "T_구매계약", fk: "발전소ID" }
      },
      groupB: {
        key: "groupB",
        label: "그룹B: 수요기업 + 판매계약 + 전기사용지",
        master: "T_수요기업",
        child: {
          table: "T_판매계약",
          fk: "수요기업ID",
          grandchild: { table: "T_전기사용지", fk: "판매계약ID" }
        }
      }
    };

    function el(tag, attrs, children) {
      var node = document.createElement(tag);
      attrs = attrs || {};
      children = children || [];

      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v === undefined || v === null) return;
        if (k === "class") node.className = v;
        else if (k === "html") node.innerHTML = v;
        else node.setAttribute(k, v);
      });

      children.forEach(function (ch) {
        if (typeof ch === "string") node.appendChild(document.createTextNode(ch));
        else if (ch) node.appendChild(ch);
      });

      return node;
    }

    // fetch() 자체가 실패하는 경우(서버에 아예 도달 못함)는 브라우저마다
    // TypeError를 던지고 메시지도 제각각입니다(크롬 "Failed to fetch",
    // 사파리 "Load failed" 등) - 이 원문 그대로가 저장/삭제 등 곳곳의
    // showToast("저장 실패: " + (e.message||e)) 패턴을 통해 사용자에게
    // 그대로 노출되고 있었습니다. 실제 원인은 대부분 (1) 실시간 입력
    // 서버가 이미 꺼진 상태(엑셀을 닫으면 자동 종료됨) 이거나 (2) 컴퓨터가
    // 절전에서 막 깨어난 직후의 아주 짧은 네트워크 끊김입니다. (2)는
    // 잠깐 뒤 한 번 더 시도하면 대부분 저절로 풀리므로 조용히 재시도하고,
    // 그래도 안 되면 (1)에 해당하는 명확한 한글 안내로 바꿔서 던집니다.
    async function fetchWithRetry(url, opts) {
      try {
        return await fetch(url, opts);
      } catch (e) {
        await new Promise(function (r) { setTimeout(r, 600); });
        try {
          return await fetch(url, opts);
        } catch (e2) {
          throw new Error(
            "서버와 연결할 수 없습니다. 엑셀의 \"대시보드생성\" 시트에서 " +
            "[실시간 입력 서버 시작] 버튼으로 서버가 켜져 있는지 확인해주세요 " +
            "(엑셀 파일을 닫으면 서버도 함께 꺼집니다)."
          );
        }
      }
    }

    async function apiGet(url) {
      var res = await fetchWithRetry(url, { cache: "no-store" });
      var text = await res.text();
      var data = {};
      try {
        data = text ? JSON.parse(text) : {};
      } catch (e) {
        throw new Error(text || "응답 파싱 실패");
      }
      if (!data.ok) throw new Error(data.error || "조회 실패");
      return data;
    }

    async function apiPost(url, payload) {
      var res = await fetchWithRetry(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload || {})
      });
      var text = await res.text();
      var data = {};
      try {
        data = text ? JSON.parse(text) : {};
      } catch (e) {
        throw new Error(text || "응답 파싱 실패");
      }
      if (!data.ok) throw new Error(data.error || "요청 실패");
      return data;
    }

    // ---------------------------------------------------------------------
    // 토스트 알림
    // ---------------------------------------------------------------------
    var toastStack = el("div", { class: "ppaf-toaststack" });

    function showToast(message, kind) {
      kind = kind || "info";
      var t = el("div", { class: "ppaf-toast ppaf-toast-" + kind }, [message]);
      toastStack.appendChild(t);
      requestAnimationFrame(function () { t.classList.add("show"); });
      setTimeout(function () {
        t.classList.remove("show");
        setTimeout(function () { t.remove(); }, 220);
      }, 3400);
    }

    // 저장/삭제 뒤 location.reload() 하면 최신 데이터를 다시 반영하기 위해
    // 페이지를 통째로 다시 불러오는데, 그냥 두면 대시보드의 state.tab이
    // 기본값('홈')으로 초기화돼 방금까지 보던 화면과 무관하게 항상 홈으로
    // 튕깁니다. reload 직전에 지금 탭을 세션스토리지에 남겨두면, 대시보드의
    // restoreTabFromSession()이 다시 로드된 뒤 그 탭으로 복원합니다.
    // 우리가 스스로 location.reload()를 부를 때(저장/삭제 뒤 최신 데이터
    // 반영)와, 사용자가 진짜로 탭/창을 닫을 때를 구분하기 위한 표시. 아래
    // pagehide 리스너가 이 값이 남아있으면 "우리가 의도한 새로고침"으로
    // 보고 서버 종료 신호를 보내지 않습니다 - 이게 없으면 저장할 때마다
    // (reload도 pagehide를 발생시키므로) 서버가 꺼져버립니다.
    function rememberDashboardTab() {
      try {
        var st = window.PPA_DASHBOARD_STATE;
        if (st && st.tab) sessionStorage.setItem("ppa_return_tab", st.tab);
        sessionStorage.setItem("ppa_intentional_reload", "1");
      } catch (e) { /* 세션스토리지 사용 불가 환경 - 조용히 무시(그냥 홈으로 감) */ }
    }

    // ---------------------------------------------------------------------
    // 필드 판별 규칙
    // ---------------------------------------------------------------------
    function isTextareaColumn(name) {
      return /비고|메모|설명|내용/.test(name);
    }
    function isDateColumn(name) {
      // "공급기한_구매"/"공급기한_판매"처럼 "기한"으로 끝나는 컬럼도 날짜
      // 입력(달력 선택 또는 직접 입력 모두 가능한 <input type=date>)으로
      // 다뤄야 하는데 예전 패턴엔 "기한"이 빠져 있어 자유 텍스트로 새고
      // 있었습니다.
      return /일자|날짜|기한|시작일|종료일|체결일|계약일/.test(name);
    }
    function isNumberColumn(name) {
      return /용량|금액|단가|비율|수량|면적|사용량|발전량|REC|MW|MWh|kW|kWh|개월|년수/.test(name);
    }
    function isRequiredColumn(tableName, name) {
      var schema = SCHEMA_BY_KEY[tableName];
      return !!(schema && (name === schema.pk || (schema.fk && Object.prototype.hasOwnProperty.call(schema.fk, name))));
    }

    // 자유 텍스트가 아니라 정해진 값 중에서만 고르게 하는 컬럼 - 컬럼 이름만
    // 보고 판단하므로 표가 달라도(발전소만 해당) 그대로 적용됩니다.
    // "현황"의 8개 값과 순서는 실제 업무 정의를 그대로 따른 것으로,
    // ppa_dashboard_render.py의 STATUS_META(아이콘·색상 매핑)와 반드시
    // 같은 8개 문구로 유지해야 합니다 - 여기서 고른 값이 그대로 저장되고
    // 그 문구로 대시보드가 아이콘/배지를 찾기 때문입니다.
    var ENUM_COLUMNS = {
      "발전원": ["태양광", "풍력", "소수력"],
      "Readiness": ["New", "Operating"],
      "현황": [
        "1. 공급 중", "2. 신고 중", "3. 상업운전 개시", "4. 공사 중",
        "5. 착공 전", "6. 이슈 발생", "7. 미확보", "99. 공급종료"
      ]
    };
    function enumOptionsFor(name) {
      return ENUM_COLUMNS[name] || null;
    }

    // MGA_Supply/MGA_Demand는 "=n/24" 수식으로 관리되던 값입니다 - 사용자가
    // 최종 소수값 대신 n만 입력하면 n/24를 계산해 그 결과를 저장합니다.
    function isFormula24Column(name) {
      return name === "MGA_Supply" || name === "MGA_Demand";
    }
    function formatFormula24Value(n) {
      var v = n / 24;
      // 부동소수점 잔여 오차 정리(소수 10자리에서 반올림) 후 끝의 불필요한 0 제거
      var s = v.toFixed(10).replace(/0+$/, "").replace(/\.$/, "");
      return s === "" || s === "-0" ? "0" : s;
    }

    function currentTable() {
      return modeSel.value;
    }

    async function getTableOptions(tableName, forceReload) {
      if (!forceReload && optionCache[tableName]) return optionCache[tableName];
      var data = await apiGet("/api/options?table=" + encodeURIComponent(tableName));
      optionCache[tableName] = data.options || [];
      return optionCache[tableName];
    }

    // /api/records 는 PK+표시용 라벨만 주는 /api/options 와 달리 전체 컬럼
    // 값을 그대로 주므로, 컬럼별 검색과 "선택 즉시 폼 채우기"에 씁니다.
    async function getTableRecords(tableName, forceReload) {
      if (!forceReload && recordCache[tableName]) return recordCache[tableName];
      var data = await apiGet("/api/records?table=" + encodeURIComponent(tableName));
      recordCache[tableName] = data.rows || [];
      return recordCache[tableName];
    }

    function recordLabel(tableName, row) {
      var schema = SCHEMA_BY_KEY[tableName];
      var pk = schema && schema.pk;
      var parts = [];
      (schema ? schema.columns : []).forEach(function (col) {
        if (col === pk) return;
        var v = (row[col] || "").toString().trim();
        if (v && parts.length < 2) parts.push(col + ": " + v);
      });
      var pkVal = pk ? row[pk] : "";
      return parts.length ? (pkVal + " — " + parts.join(", ")) : String(pkVal || "");
    }

    // 컬럼 선택(전체/특정 컬럼) + 키워드 입력으로 기존 데이터를 찾는 재사용
    // 가능한 검색 위젯. onSelect(row) 는 사용자가 결과 목록에서 하나를
    // 클릭했을 때 호출됩니다. recordCache[tableName] 이 미리 채워져 있어야
    // 합니다(호출 전에 getTableRecords 로 로드).
    function buildPicker(tableName, onSelect) {
      var wrap = el("div", { class: "ppaf-picker" });
      var controls = el("div", { class: "ppaf-picker-controls" });
      var colSel = el("select", { class: "ppaf-input ppaf-picker-col" });
      colSel.appendChild(el("option", { value: "" }, ["전체 컬럼"]));
      var schema = SCHEMA_BY_KEY[tableName];
      (schema ? schema.columns : []).forEach(function (col) {
        colSel.appendChild(el("option", { value: col }, [col]));
      });
      var qInput = el("input", {
        class: "ppaf-input ppaf-picker-q",
        type: "text",
        placeholder: "검색어 입력 (예: 이름·ID 일부...)"
      });
      var results = el("div", { class: "ppaf-picker-results" });

      controls.appendChild(colSel);
      controls.appendChild(qInput);
      wrap.appendChild(controls);
      wrap.appendChild(results);

      function runFilter() {
        var q = (qInput.value || "").trim().toLowerCase();
        var col = colSel.value;
        var rows = recordCache[tableName] || [];
        results.innerHTML = "";
        if (!q) {
          results.appendChild(el("div", { class: "ppaf-picker-hint" }, ["검색어를 입력하면 목록이 나타납니다. (총 " + rows.length + "건)"]));
          return;
        }
        var matched = rows.filter(function (row) {
          if (col) return String(row[col] || "").toLowerCase().indexOf(q) !== -1;
          return Object.keys(row).some(function (k) {
            return String(row[k] || "").toLowerCase().indexOf(q) !== -1;
          });
        });
        if (!matched.length) {
          results.appendChild(el("div", { class: "ppaf-picker-hint" }, ["일치하는 데이터가 없습니다."]));
          return;
        }
        var shown = matched.slice(0, 30);
        shown.forEach(function (row) {
          var item = el("button", { class: "ppaf-picker-item", type: "button" }, [recordLabel(tableName, row)]);
          item.addEventListener("click", function () { onSelect(row); });
          results.appendChild(item);
        });
        if (matched.length > shown.length) {
          results.appendChild(
            el("div", { class: "ppaf-picker-more" }, [(matched.length - shown.length) + "건 더 있음 - 검색어를 구체화해주세요."])
          );
        }
      }

      qInput.addEventListener("input", runFilter);
      colSel.addEventListener("change", runFilter);
      runFilter();

      return wrap;
    }

    async function prepareOptionsForTable(tableName, forceReload) {
      await getTableOptions(tableName, forceReload);
      var refs = FIELD_LOOKUP[tableName] || {};
      var refTables = Object.keys(refs).map(function (k) { return refs[k]; });
      for (var i = 0; i < refTables.length; i++) {
        await getTableOptions(refTables[i], forceReload);
      }
    }

    function getLookupTable(tableName, columnName) {
      var refs = FIELD_LOOKUP[tableName] || {};
      return refs[columnName] || "";
    }

    function buildSelect(fieldName, options) {
      var input = el("select", { class: "ppaf-input", "data-name": fieldName });
      input.appendChild(el("option", { value: "" }, ["선택"]));
      (options || []).forEach(function (opt) {
        input.appendChild(el("option", { value: opt.value }, [opt.label]));
      });
      return input;
    }

    // ---------------------------------------------------------------------
    // 실시간 유효성 검사
    // ---------------------------------------------------------------------
    function markInvalid(wrap, input, msg) {
      wrap.classList.add("invalid");
      var note = wrap.querySelector(".ppaf-fieldnote");
      if (!note) {
        note = el("div", { class: "ppaf-fieldnote ppaf-fieldnote-error" });
        wrap.appendChild(note);
      }
      note.className = "ppaf-fieldnote ppaf-fieldnote-error";
      note.textContent = msg;
    }

    function markHint(wrap, msg) {
      wrap.classList.remove("invalid");
      var note = wrap.querySelector(".ppaf-fieldnote");
      if (!note) {
        note = el("div", { class: "ppaf-fieldnote" });
        wrap.appendChild(note);
      }
      note.className = "ppaf-fieldnote ppaf-fieldnote-hint";
      note.textContent = msg;
    }

    function clearFieldNote(wrap) {
      wrap.classList.remove("invalid");
      var note = wrap.querySelector(".ppaf-fieldnote");
      if (note) note.remove();
    }

    // alreadyLoaded: 이 레코드가 "기존 데이터를 불러온" 상태인지(true면 PK
    // 중복 힌트를 안 보여줌 - 자기 자신이니까). 단일입력에서는
    // formState.loadedPk 를, 그룹입력에서는 각 항목의 existing 플래그를 씁니다.
    function validateFieldLive(tableName, columnName, wrap, input, alreadyLoaded) {
      var schema = SCHEMA_BY_KEY[tableName];
      var required = isRequiredColumn(tableName, columnName);
      var value = (input.value || "").trim();

      if (required && !value) {
        markInvalid(wrap, input, columnName + "는 필수입니다.");
        return false;
      }

      // 새 데이터 입력 중인데 PK가 이미 존재하는 값이면 - 막지는 않되(저장하면
      // 그 데이터를 덮어써 수정하는 것과 같으므로) 알려줍니다.
      if (schema && columnName === schema.pk && value && !alreadyLoaded) {
        var exists = (optionCache[tableName] || []).some(function (o) { return o.value === value; });
        if (exists) {
          markHint(wrap, "이미 존재하는 " + columnName + "입니다 - 저장하면 그 데이터가 수정됩니다.");
          return true;
        }
      }

      clearFieldNote(wrap);
      return true;
    }

    function validateAllFields(tableName) {
      var ok = true;
      fields.querySelectorAll(".ppaf-row").forEach(function (wrap) {
        var input = wrap.querySelector("[data-name^='fld_']");
        if (!input) return;
        var columnName = input.getAttribute("data-name").replace("fld_", "");
        if (!validateFieldLive(tableName, columnName, wrap, input, !!formState.loadedPk)) ok = false;
      });
      return ok;
    }

    // 필수 컬럼이 채워져 있는지만 확인하는 가벼운 버전 - 그룹 저장 직전에
    // DOM이 아니라 in-memory record 객체 자체를 검사할 때 씁니다.
    function validateRecordRequired(tableName, record) {
      var schema = SCHEMA_BY_KEY[tableName];
      var errs = [];
      (schema ? schema.columns : []).forEach(function (col) {
        if (isRequiredColumn(tableName, col) && !(record[col] || "").toString().trim()) {
          errs.push(col + "는 필수입니다.");
        }
      });
      return errs;
    }

    // MGA_Supply/MGA_Demand 전용: 화면엔 "n" 입력칸만 두고, 실제 저장 값
    // (n/24)은 숨김 input(data-name 보유)에 넣어 collectRecord/fillRecord가
    // 그대로 다루게 합니다. wrap에 .ppaf-formula24 클래스를 붙여 fillRecord가
    // 로드 시 n 표시칸을 역산해 채울 수 있게 합니다.
    function createFormula24Field(tableName, columnName) {
      var wrap = el("div", { class: "ppaf-row ppaf-formula24" });
      var labelClass = isRequiredColumn(tableName, columnName) ? "ppaf-label ppaf-required" : "ppaf-label";
      wrap.appendChild(el("label", { class: labelClass }, [columnName + " (= n / 24)"]));

      var fieldName = "fld_" + columnName;
      var nInput = el("input", { class: "ppaf-input", type: "number", step: "any", placeholder: "n 입력 (예: 3.2)" });
      var hidden = el("input", { type: "hidden", "data-name": fieldName });
      var display = el("div", { class: "ppaf-formula24-display" }, ["n을 입력하면 자동 계산됩니다."]);

      function recompute() {
        var n = parseFloat(nInput.value);
        if (nInput.value === "" || isNaN(n)) {
          hidden.value = "";
          display.textContent = "n을 입력하면 자동 계산됩니다.";
          return;
        }
        var v = formatFormula24Value(n);
        hidden.value = v;
        display.textContent = "= " + n + " / 24 = " + v;
      }

      wrap.appendChild(nInput);
      wrap.appendChild(display);
      wrap.appendChild(hidden);

      var onLiveCheck = function () {
        recompute();
        validateFieldLive(tableName, columnName, wrap, hidden, !!formState.loadedPk);
      };
      nInput.addEventListener("input", onLiveCheck);
      nInput.addEventListener("blur", onLiveCheck);

      return wrap;
    }

    // 로드된 레코드의 저장값(n/24 계산 결과)로부터 n 표시칸을 역산해 채웁니다.
    // fillRecord()가 hidden input에 값을 채운 직후 호출됩니다.
    function syncFormula24Displays(container) {
      container.querySelectorAll(".ppaf-formula24").forEach(function (wrap) {
        var hidden = wrap.querySelector("input[type=hidden]");
        var nInput = wrap.querySelector("input[type=number]");
        var display = wrap.querySelector(".ppaf-formula24-display");
        if (!hidden || !nInput) return;
        var v = parseFloat(hidden.value);
        if (hidden.value === "" || isNaN(v)) {
          nInput.value = "";
          if (display) display.textContent = "n을 입력하면 자동 계산됩니다.";
          return;
        }
        var n = Math.round(v * 24 * 1000) / 1000;
        nInput.value = String(n);
        if (display) display.textContent = "= " + n + " / 24 = " + hidden.value;
      });
    }

    function createField(tableName, columnName) {
      if (isFormula24Column(columnName)) return createFormula24Field(tableName, columnName);

      var wrap = el("div", { class: "ppaf-row" });
      var labelClass = isRequiredColumn(tableName, columnName) ? "ppaf-label ppaf-required" : "ppaf-label";
      var label = el("label", { class: labelClass }, [columnName]);
      var fieldName = "fld_" + columnName;
      var input;

      var lookupTable = getLookupTable(tableName, columnName);
      var enumOptions = enumOptionsFor(columnName);
      if (lookupTable && (optionCache[lookupTable] || []).length > 0) {
        input = buildSelect(fieldName, optionCache[lookupTable] || []);
      } else if (enumOptions) {
        input = buildSelect(fieldName, enumOptions.map(function (v) { return { value: v, label: v }; }));
      } else if (isTextareaColumn(columnName)) {
        input = el("textarea", { class: "ppaf-textarea", "data-name": fieldName });
      } else if (isDateColumn(columnName)) {
        input = el("input", { class: "ppaf-input", type: "date", "data-name": fieldName });
      } else if (isNumberColumn(columnName)) {
        input = el("input", { class: "ppaf-input", type: "number", step: "any", "data-name": fieldName });
      } else {
        input = el("input", { class: "ppaf-input", type: "text", "data-name": fieldName });
      }

      wrap.appendChild(label);
      wrap.appendChild(input);

      var onLiveCheck = function () { validateFieldLive(tableName, columnName, wrap, input, !!formState.loadedPk); };
      input.addEventListener("blur", onLiveCheck);
      input.addEventListener("change", onLiveCheck);
      input.addEventListener("input", function () {
        if (wrap.classList.contains("invalid")) onLiveCheck();
      });

      return wrap;
    }

    // ---------------------------------------------------------------------
    // 그룹(마스터+자식) 폼에서 쓰는 필드 생성기 - createField와 달리
    // data-name 기반 collectRecord() 를 쓰지 않고, 입력값을 바로 넘겨받은
    // record 객체에 반영합니다(카드가 여러 개 동시에 떠 있으므로).
    // ---------------------------------------------------------------------
    function buildFieldsGrid(tableName, record, opts) {
      opts = opts || {};
      var exclude = opts.exclude || [];
      var alreadyLoaded = !!opts.alreadyLoaded;
      var grid = el("div", { class: "ppaf-fields" });
      var schema = SCHEMA_BY_KEY[tableName];
      var columns = (schema && schema.columns) || [];
      columns.forEach(function (col) {
        if (exclude.indexOf(col) !== -1) return;
        grid.appendChild(buildBoundFieldRow(tableName, col, record, alreadyLoaded));
      });
      return grid;
    }

    function buildBoundFieldRow(tableName, columnName, record, alreadyLoaded) {
      if (isFormula24Column(columnName)) return buildFormula24BoundRow(tableName, columnName, record, alreadyLoaded);

      var wrap = el("div", { class: "ppaf-row" });
      var labelClass = isRequiredColumn(tableName, columnName) ? "ppaf-label ppaf-required" : "ppaf-label";
      wrap.appendChild(el("label", { class: labelClass }, [columnName]));

      var input;
      var lookupTable = getLookupTable(tableName, columnName);
      var enumOptions = enumOptionsFor(columnName);
      var currentVal = record[columnName] || "";
      if (lookupTable && (optionCache[lookupTable] || []).length > 0) {
        input = buildSelect("x", optionCache[lookupTable] || []);
      } else if (enumOptions) {
        input = buildSelect("x", enumOptions.map(function (v) { return { value: v, label: v }; }));
      } else if (isTextareaColumn(columnName)) {
        input = el("textarea", { class: "ppaf-textarea" });
      } else if (isDateColumn(columnName)) {
        input = el("input", { class: "ppaf-input", type: "date" });
      } else if (isNumberColumn(columnName)) {
        input = el("input", { class: "ppaf-input", type: "number", step: "any" });
      } else {
        input = el("input", { class: "ppaf-input", type: "text" });
      }
      input.removeAttribute("data-name");
      if (input.tagName === "SELECT" && currentVal &&
          !Array.prototype.some.call(input.options, function (o) { return o.value === currentVal; })) {
        input.appendChild(el("option", { value: currentVal }, [currentVal + " (목록에 없는 기존 값)"]));
      }
      input.value = currentVal;
      wrap.appendChild(input);

      var sync = function () { record[columnName] = input.value || ""; };
      var onLiveCheck = function () {
        sync();
        validateFieldLive(tableName, columnName, wrap, input, alreadyLoaded);
      };
      input.addEventListener("blur", onLiveCheck);
      input.addEventListener("change", onLiveCheck);
      input.addEventListener("input", function () {
        sync();
        if (wrap.classList.contains("invalid")) onLiveCheck();
      });

      return wrap;
    }

    function buildFormula24BoundRow(tableName, columnName, record, alreadyLoaded) {
      var wrap = el("div", { class: "ppaf-row ppaf-formula24" });
      var labelClass = isRequiredColumn(tableName, columnName) ? "ppaf-label ppaf-required" : "ppaf-label";
      wrap.appendChild(el("label", { class: labelClass }, [columnName + " (= n / 24)"]));

      var nInput = el("input", { class: "ppaf-input", type: "number", step: "any", placeholder: "n 입력 (예: 3.2)" });
      var display = el("div", { class: "ppaf-formula24-display" });

      function renderDisplay(v) {
        var n = parseFloat(v);
        display.textContent = (v === "" || v === null || v === undefined || isNaN(n))
          ? "n을 입력하면 자동 계산됩니다."
          : "저장값: " + v;
      }

      var existingNum = parseFloat(record[columnName]);
      if (!isNaN(existingNum)) nInput.value = String(Math.round(existingNum * 24 * 1000) / 1000);
      renderDisplay(record[columnName]);

      function recompute() {
        if (nInput.value === "" || isNaN(parseFloat(nInput.value))) {
          record[columnName] = "";
          renderDisplay("");
          return;
        }
        var v = formatFormula24Value(parseFloat(nInput.value));
        record[columnName] = v;
        renderDisplay(v);
      }
      var onLiveCheck = function () {
        recompute();
        validateFieldLive(tableName, columnName, wrap, nInput, alreadyLoaded);
      };
      nInput.addEventListener("input", onLiveCheck);
      nInput.addEventListener("blur", onLiveCheck);

      wrap.appendChild(nInput);
      wrap.appendChild(display);
      return wrap;
    }

    function collectRecord() {
      var record = {};
      fields.querySelectorAll("[data-name^='fld_']").forEach(function (inp) {
        var key = inp.getAttribute("data-name").replace("fld_", "");
        record[key] = inp.value || "";
      });
      return record;
    }

    function fillRecord(record) {
      record = record || {};
      fields.querySelectorAll("[data-name^='fld_']").forEach(function (inp) {
        var key = inp.getAttribute("data-name").replace("fld_", "");
        var val = record[key] || "";
        // 목록형(select)인데 지금 값이 옵션 목록에 없으면(예: 예전에 자유
        // 입력하던 시절의 값) 조용히 지우지 않고 임시 옵션으로 보존합니다.
        if (inp.tagName === "SELECT" && val &&
            !Array.prototype.some.call(inp.options, function (o) { return o.value === val; })) {
          inp.appendChild(el("option", { value: val }, [val + " (목록에 없는 기존 값)"]));
        }
        inp.value = val;
      });
      fields.querySelectorAll(".ppaf-row").forEach(clearFieldNote);
      syncFormula24Displays(fields);
    }

    function clearRecord() {
      fields.querySelectorAll("[data-name^='fld_']").forEach(function (inp) { inp.value = ""; });
      fields.querySelectorAll(".ppaf-row").forEach(clearFieldNote);
      syncFormula24Displays(fields);
      formState.loadedPk = null;
      formDirty = false;
      updateDeleteButtonState();
      var pq = fields.querySelector(".ppaf-picker-q");
      if (pq) { pq.value = ""; pq.dispatchEvent(new Event("input")); }
      var loadedIndicator = fields.querySelector(".ppaf-loaded-indicator");
      if (loadedIndicator) loadedIndicator.style.display = "none";
      var dupBtn = fields.querySelector(".ppaf-dup-btn");
      if (dupBtn) dupBtn.disabled = true;
    }

    function updateDeleteButtonState() {
      if (appMode !== "single") return;
      deleteBtn.disabled = !formState.loadedPk;
      deleteBtn.title = formState.loadedPk ? "" : "먼저 검색으로 기존 데이터를 선택하세요.";
    }

    // 단일입력 화면에 레코드를 불러온 상태로 반영합니다(필드 채우기 + 불러옴
    // 표시 칩 + 삭제/복제 버튼 활성화). 컬럼별 검색 결과 클릭과, 외부(표 탭의
    // 상세 모달 "수정"/"삭제" 버튼)에서 레코드를 여는 경로 둘 다 이 함수를
    // 씁니다 - closure가 아니라 클래스로 요소를 찾으므로 호출 시점의 toolbar
    // DOM이 어떻게 만들어졌는지와 무관하게 항상 동작합니다.
    function applyLoadedRecord(tableName, row) {
      fillRecord(row || {});
      var schema = SCHEMA_BY_KEY[tableName];
      formState.loadedPk = schema ? (row[schema.pk] || null) : null;
      formDirty = false;
      updateDeleteButtonState();
      var indicator = fields.querySelector(".ppaf-loaded-indicator");
      var label = fields.querySelector(".ppaf-loaded-label");
      var dupBtn = fields.querySelector(".ppaf-dup-btn");
      if (indicator) indicator.style.display = "";
      if (label) label.textContent = recordLabel(tableName, row);
      if (dupBtn) dupBtn.disabled = false;
    }

    function renderToolbar(tableName, autofocusPicker) {
      var toolbar = el("div", { class: "ppaf-toolbar-wrap" });
      var newBtn = el("button", { class: "ppaf-btn", type: "button" }, ["새 입력"]);
      var dupBtn = el("button", { class: "ppaf-btn ppaf-dup-btn", type: "button", disabled: "disabled" }, ["복제해서 새로 입력"]);

      var loadedIndicator = el("div", { class: "ppaf-loaded-indicator", style: "display:none" });
      var loadedLabel = el("span", { class: "ppaf-loaded-label" });
      var loadedClearBtn = el("button", { class: "ppaf-chip-x", type: "button", title: "선택 해제" }, ["✕"]);
      loadedIndicator.appendChild(el("span", { class: "ppaf-loaded-dot" }));
      loadedIndicator.appendChild(loadedLabel);
      loadedIndicator.appendChild(loadedClearBtn);

      newBtn.addEventListener("click", function () {
        if (!confirmDiscardIfDirty()) return;
        clearRecord();
      });
      loadedClearBtn.addEventListener("click", function () {
        if (!confirmDiscardIfDirty()) return;
        clearRecord();
      });

      dupBtn.addEventListener("click", function () {
        if (!formState.loadedPk) return;
        var record = collectRecord();
        var schema = SCHEMA_BY_KEY[tableName];
        if (schema) record[schema.pk] = "";
        fillRecord(record);
        formState.loadedPk = null;
        formDirty = true;
        updateDeleteButtonState();
        loadedIndicator.style.display = "none";
        dupBtn.disabled = true;
        var pkInput = fields.querySelector("[data-name='fld_" + (schema ? schema.pk : "") + "']");
        if (pkInput) pkInput.focus();
        showToast("복제되었습니다 - 새 " + (schema ? schema.pk : "PK") + "를 입력하고 저장하세요.", "info");
      });

      var pickerLabel = el("div", { class: "ppaf-picker-label" }, ["기존 데이터 불러오기 (컬럼별 검색)"]);
      var picker = buildPicker(tableName, function (row) {
        if (!confirmDiscardIfDirty()) return;
        applyLoadedRecord(tableName, row);
        showToast("불러왔습니다.", "info");
      });

      toolbar.appendChild(newBtn);
      toolbar.appendChild(dupBtn);
      toolbar.appendChild(loadedIndicator);
      toolbar.appendChild(pickerLabel);
      toolbar.appendChild(picker);

      if (tableName === "T_수급매칭") {
        var matchWidget = buildSupplyMatchWidget();
        toolbar.appendChild(matchWidget.trigger);
        fields.appendChild(toolbar);
        fields.appendChild(matchWidget.panel);
      } else {
        fields.appendChild(toolbar);
      }

      if (autofocusPicker) {
        var q = picker.querySelector(".ppaf-picker-q");
        if (q) setTimeout(function () { q.focus(); }, 30);
      }
    }

    // -----------------------------------------------------------------------
    // 수급매칭ID처럼 "접두어 + 숫자" 규칙을 쓰는 PK의 다음 값을 실제 데이터에서
    // 그대로 추정합니다 - 데모에서 정한 규칙이 아니라 지금 엑셀에 들어있는
    // 값들의 접두어/자리수를 그때그때 관찰해서 만들므로, 실제 업무 명명 규칙이
    // 무엇이든(예: "매칭-001", "SM-2024-01" 등) 그대로 따라갑니다. 숫자로
    // 끝나는 값이 하나도 없으면 빈 문자열을 돌려줘 사용자가 직접 입력하게
    // 둡니다(잘못 추측해 강제로 채우지 않음).
    function suggestNextId(existingIds) {
      var groups = {};
      (existingIds || []).forEach(function (id) {
        var m = /^(.*?)(\d+)$/.exec(String(id || ""));
        if (!m) return;
        var prefix = m[1], width = m[2].length, num = parseInt(m[2], 10);
        var g = groups[prefix];
        if (!g) g = groups[prefix] = { count: 0, max: -1, width: width };
        g.count++;
        if (num > g.max) g.max = num;
        g.width = Math.max(g.width, width);
      });
      var bestPrefix = null, bestCount = -1;
      Object.keys(groups).forEach(function (p) {
        if (groups[p].count > bestCount) { bestCount = groups[p].count; bestPrefix = p; }
      });
      if (bestPrefix === null) return "";
      var g = groups[bestPrefix];
      var numStr = String(g.max + 1);
      while (numStr.length < g.width) numStr = "0" + numStr;
      return bestPrefix + numStr;
    }

    // -----------------------------------------------------------------------
    // 수급매칭 전용 "카드로 쉽게 매칭하기" - 구매계약 카드 하나 + 전기사용지
    // 카드 하나를 클릭으로 골라 연결하면, 아래 일반 입력 폼의 전기사용지ID/
    // 구매계약ID 칸을 그대로 채워줍니다(수급매칭ID도 비어있으면 자동 제안).
    // 저장/검증은 기존 saveCurrentTable() 흐름을 그대로 씁니다 - 새 API 없음.
    // -----------------------------------------------------------------------
    function buildSupplyMatchWidget() {
      var trigger = el("button", { class: "ppaf-btn ppaf-matchtrigger", type: "button" }, ["🔗 카드로 쉽게 매칭하기"]);
      var panel = el("div", { class: "ppaf-matchwrap", style: "display:none" });
      var loaded = false;
      var sel = { 구매: null, 전기: null };
      var existingPairs = {}; // "전기사용지ID||구매계약ID" -> true
      var joinPlantName = {}; // 발전소ID -> 발전소명

      function pairKey(feId, pcId) { return String(feId) + "||" + String(pcId); }

      function recomputeExistingPairs() {
        existingPairs = {};
        (recordCache["T_수급매칭"] || []).forEach(function (r) {
          existingPairs[pairKey(r["전기사용지ID"], r["구매계약ID"])] = true;
        });
      }

      function matchCountFor(col, idVal) {
        return (recordCache["T_수급매칭"] || []).filter(function (r) { return String(r[col] || "") === String(idVal); }).length;
      }

      var pcSearch = el("input", { class: "ppaf-input ppaf-matchsearch", type: "text", placeholder: "구매계약 검색 (ID·발전소명 등)" });
      var feSearch = el("input", { class: "ppaf-input ppaf-matchsearch", type: "text", placeholder: "전기사용지 검색 (ID·이름 등)" });
      var pcList = el("div", { class: "ppaf-matchlist" });
      var feList = el("div", { class: "ppaf-matchlist" });
      var previewText = el("div", { class: "ppaf-matchpreview" }, ["구매계약과 전기사용지를 각각 하나씩 선택하세요."]);
      var connectBtn = el("button", { class: "ppaf-btn primary", type: "button", disabled: "disabled" }, ["연결하기"]);
      var closeBtn2 = el("button", { class: "ppaf-btn", type: "button" }, ["닫기"]);

      function cardRow(kind, row, schema) {
        var idVal = row[schema.pk];
        var isSel = sel[kind] && sel[kind][schema.pk] === idVal;
        var card = el("div", { class: "ppaf-matchcard" + (isSel ? " selected" : "") });
        var count = matchCountFor(kind === "구매" ? "구매계약ID" : "전기사용지ID", idVal);
        var pill = el("span", { class: "ppaf-matchpill" + (count > 0 ? " on" : "") }, [count > 0 ? count + "건 매칭됨" : "미매칭"]);
        card.appendChild(el("div", { class: "ppaf-matchcard-head" }, [el("b", {}, [idVal]), pill]));
        if (kind === "구매") {
          var plant = joinPlantName[row["발전소ID"]] || row["발전소ID"] || "";
          card.appendChild(el("div", { class: "ppaf-matchcard-line" }, [plant + " · " + (row["구매계약용량(MW)"] || "?") + " MW"]));
          card.appendChild(el("div", { class: "ppaf-matchcard-sub" }, ["담당자: " + (row["구매 담당자"] || "-") + " · 공급기한: " + (row["공급기한_구매"] || "-")]));
        } else {
          card.appendChild(el("div", { class: "ppaf-matchcard-line" }, [(row["전기사용지명"] || "") + " · " + (row["전기사용지계약용량(MW)"] || "?") + " MW"]));
          card.appendChild(el("div", { class: "ppaf-matchcard-sub" }, ["판매계약: " + (row["판매계약ID"] || "-")]));
        }
        card.addEventListener("click", function () {
          sel[kind] = isSel ? null : row;
          renderLists();
          renderPreview();
        });
        return card;
      }

      function filteredRows(tableName, query) {
        var rows = recordCache[tableName] || [];
        var q = (query || "").trim().toLowerCase();
        if (!q) return rows;
        return rows.filter(function (r) {
          return Object.keys(r).some(function (k) { return String(r[k] || "").toLowerCase().indexOf(q) !== -1; });
        });
      }

      function renderLists() {
        var pcSchema = SCHEMA_BY_KEY["T_구매계약"], feSchema = SCHEMA_BY_KEY["T_전기사용지"];
        pcList.innerHTML = "";
        filteredRows("T_구매계약", pcSearch.value).forEach(function (r) { pcList.appendChild(cardRow("구매", r, pcSchema)); });
        if (!pcList.children.length) pcList.appendChild(el("div", { class: "ppaf-picker-hint" }, ["일치하는 구매계약이 없습니다."]));
        feList.innerHTML = "";
        filteredRows("T_전기사용지", feSearch.value).forEach(function (r) { feList.appendChild(cardRow("전기", r, feSchema)); });
        if (!feList.children.length) feList.appendChild(el("div", { class: "ppaf-picker-hint" }, ["일치하는 전기사용지가 없습니다."]));
      }

      function renderPreview() {
        if (!sel.구매 || !sel.전기) {
          previewText.className = "ppaf-matchpreview";
          previewText.textContent = "구매계약과 전기사용지를 각각 하나씩 선택하세요.";
          connectBtn.disabled = true;
          return;
        }
        var pcId = sel.구매["구매계약ID"], feId = sel.전기["전기사용지ID"];
        if (existingPairs[pairKey(feId, pcId)]) {
          previewText.className = "ppaf-matchpreview warn";
          previewText.textContent = "⚠ 이미 이 조합으로 매칭된 수급매칭이 있습니다 - 다른 조합을 선택해주세요.";
          connectBtn.disabled = true;
          return;
        }
        previewText.className = "ppaf-matchpreview ok";
        previewText.textContent = "구매계약 " + pcId + "  ↔  전기사용지 " + feId;
        connectBtn.disabled = false;
      }

      pcSearch.addEventListener("input", renderLists);
      feSearch.addEventListener("input", renderLists);

      connectBtn.addEventListener("click", function () {
        var feInput = fields.querySelector('[data-name="fld_전기사용지ID"]');
        var pcInput = fields.querySelector('[data-name="fld_구매계약ID"]');
        var pkSchema = SCHEMA_BY_KEY["T_수급매칭"];
        var pkInput = pkSchema ? fields.querySelector('[data-name="fld_' + pkSchema.pk + '"]') : null;

        if (feInput) { feInput.value = sel.전기["전기사용지ID"]; feInput.dispatchEvent(new Event("change", { bubbles: true })); }
        if (pcInput) { pcInput.value = sel.구매["구매계약ID"]; pcInput.dispatchEvent(new Event("change", { bubbles: true })); }
        if (pkInput && !pkInput.value && !formState.loadedPk) {
          var suggested = suggestNextId((recordCache["T_수급매칭"] || []).map(function (r) { return r[pkSchema.pk]; }));
          if (suggested) { pkInput.value = suggested; pkInput.dispatchEvent(new Event("change", { bubbles: true })); }
        }
        formDirty = true;

        panel.style.display = "none";
        showToast("선택한 구매계약·전기사용지가 연결되었습니다. 현황을 선택하고 저장하세요.", "success");
        var statusInput = fields.querySelector('[data-name="fld_현황"]');
        if (statusInput) { statusInput.focus(); statusInput.scrollIntoView({ block: "center", behavior: "smooth" }); }
      });

      closeBtn2.addEventListener("click", function () { panel.style.display = "none"; });

      var pcCol = el("div", { class: "ppaf-matchcol" }, [
        el("div", { class: "ppaf-matchcol-title" }, ["구매계약 선택"]), pcSearch, pcList
      ]);
      var feCol = el("div", { class: "ppaf-matchcol" }, [
        el("div", { class: "ppaf-matchcol-title" }, ["전기사용지 선택"]), feSearch, feList
      ]);
      var grid = el("div", { class: "ppaf-matchgrid" }, [pcCol, feCol]);
      var foot = el("div", { class: "ppaf-matchfoot" }, [previewText, el("div", { class: "ppaf-matchfoot-btns" }, [closeBtn2, connectBtn])]);
      panel.appendChild(grid);
      panel.appendChild(foot);

      trigger.addEventListener("click", function () {
        if (panel.style.display !== "none") { panel.style.display = "none"; return; }
        panel.style.display = "";
        (async function () {
          if (!loaded) {
            trigger.disabled = true;
            try {
              await Promise.all([
                getTableRecords("T_구매계약", false),
                getTableRecords("T_전기사용지", false),
                getTableRecords("T_발전소", false)
              ]);
              (recordCache["T_발전소"] || []).forEach(function (p) { joinPlantName[p["발전소ID"]] = p["발전소명"]; });
              loaded = true;
            } catch (e) {
              showToast("매칭용 데이터 불러오기 실패: " + (e.message || e), "error");
            } finally {
              trigger.disabled = false;
            }
          }
          recomputeExistingPairs();
          sel = { 구매: null, 전기: null };
          renderLists();
          renderPreview();
        })();
      });

      return { trigger: trigger, panel: panel };
    }

    function renderFields(tableName) {
      var schema = SCHEMA_BY_KEY[tableName];
      var columns = (schema && schema.columns) || [];
      columns.forEach(function (col) {
        fields.appendChild(createField(tableName, col));
      });
    }

    // ---------------------------------------------------------------------
    // 그룹(PK 연계 표 일괄) CRUD - 마스터 1건 + 자식(+손자) N건을 한 화면에서
    // 편집하고 /api/batch 로 한 번에 반영합니다. 개별입력 화면과 별개로
    // groupWrap 안에서 그때그때 다시 그립니다(state 기반 재렌더 방식).
    // ---------------------------------------------------------------------
    function newGroupState(kind) {
      return { kind: kind, master: { record: {}, existing: false }, children: [] };
    }

    async function loadChildrenFor(def, masterPkVal) {
      var childDef = def.child;
      var childSchema = SCHEMA_BY_KEY[childDef.table];
      var rows = await getTableRecords(childDef.table, true);
      var matched = rows.filter(function (r) { return String(r[childDef.fk] || "") === String(masterPkVal); });

      var children = [];
      for (var i = 0; i < matched.length; i++) {
        var entry = { table: childDef.table, fk: childDef.fk, record: matched[i], existing: true, deleted: false, grand: [] };
        if (childDef.grandchild) {
          var gdef = childDef.grandchild;
          var childPkVal = matched[i][childSchema.pk];
          var grows = await getTableRecords(gdef.table, true);
          var gmatched = grows.filter(function (r) { return String(r[gdef.fk] || "") === String(childPkVal); });
          entry.grand = gmatched.map(function (g) {
            return { table: gdef.table, fk: gdef.fk, record: g, existing: true, deleted: false };
          });
        }
        children.push(entry);
      }
      return children;
    }

    function groupLabel(tableName) {
      return TABLE_META[tableName] ? TABLE_META[tableName].label : tableName;
    }

    function buildGroupMasterToolbar(def) {
      var wrap = el("div", { class: "ppaf-toolbar-wrap" });
      var newBtn = el("button", { class: "ppaf-btn", type: "button" }, ["새 " + groupLabel(def.master) + " 입력"]);

      var loadedIndicator = el("div", { class: "ppaf-loaded-indicator", style: groupState.master.existing ? "" : "display:none" });
      var loadedLabel = el("span", { class: "ppaf-loaded-label" }, [groupState.master.existing ? recordLabel(def.master, groupState.master.record) : ""]);
      var loadedClearBtn = el("button", { class: "ppaf-chip-x", type: "button", title: "선택 해제" }, ["✕"]);
      loadedIndicator.appendChild(el("span", { class: "ppaf-loaded-dot" }));
      loadedIndicator.appendChild(loadedLabel);
      loadedIndicator.appendChild(loadedClearBtn);

      newBtn.addEventListener("click", function () {
        if (!confirmDiscardIfDirty()) return;
        groupDirty = false;
        groupState = newGroupState(def.key);
        renderGroup().catch(console.error);
      });
      loadedClearBtn.addEventListener("click", function () {
        if (!confirmDiscardIfDirty()) return;
        groupDirty = false;
        groupState = newGroupState(def.key);
        renderGroup().catch(console.error);
      });

      var pickerLabel = el("div", { class: "ppaf-picker-label" }, ["기존 " + groupLabel(def.master) + " 불러오기 (컬럼별 검색)"]);
      var picker = buildPicker(def.master, function (row) {
        if (!confirmDiscardIfDirty()) return;
        var schema = SCHEMA_BY_KEY[def.master];
        var pkVal = schema ? row[schema.pk] : "";
        groupState.master.record = row;
        groupState.master.existing = true;
        loadChildrenFor(def, pkVal)
          .then(function (children) {
            groupState.children = children;
            groupDirty = false;
            return renderGroup();
          })
          .then(function () {
            showToast("불러왔습니다 (하위 " + groupState.children.length + "건 포함).", "info");
          })
          .catch(function (e) {
            showToast("하위 데이터 불러오기 실패: " + (e.message || e), "error");
          });
      });

      wrap.appendChild(newBtn);
      wrap.appendChild(loadedIndicator);
      wrap.appendChild(pickerLabel);
      wrap.appendChild(picker);
      return wrap;
    }

    function buildChildCard(def, child) {
      var childDef = def.child;
      var card = el("div", { class: "ppaf-childcard" + (child.deleted ? " deleted" : "") });
      var badge = el("span", { class: "ppaf-badge " + (child.existing ? "existing" : "new") }, [child.existing ? "기존" : "신규"]);

      var toggleBtn;
      if (child.existing) {
        toggleBtn = el("button", { class: "ppaf-btn danger", type: "button" }, [child.deleted ? "삭제 취소" : "삭제 표시"]);
        toggleBtn.addEventListener("click", function () {
          child.deleted = !child.deleted;
          if (child.deleted && child.grand) {
            child.grand = child.grand.filter(function (g) { return g.existing; });
            child.grand.forEach(function (g) { g.deleted = true; });
          }
          groupDirty = true;
          renderGroup().catch(console.error);
        });
      } else {
        toggleBtn = el("button", { class: "ppaf-btn danger", type: "button" }, ["제거"]);
        toggleBtn.addEventListener("click", function () {
          var i = groupState.children.indexOf(child);
          if (i !== -1) groupState.children.splice(i, 1);
          groupDirty = true;
          renderGroup().catch(console.error);
        });
      }

      card.appendChild(el("div", { class: "ppaf-childcard-head" }, [badge, toggleBtn]));

      if (!child.deleted) {
        card.appendChild(buildFieldsGrid(childDef.table, child.record, { exclude: [childDef.fk], alreadyLoaded: child.existing }));
        if (childDef.grandchild) {
          card.appendChild(buildGrandchildSection(childDef, child));
        }
      } else {
        var note = "이 항목은 저장 시 삭제됩니다.";
        if (childDef.grandchild && child.grand && child.grand.length) {
          note = "이 항목은 저장 시 하위 " + groupLabel(childDef.grandchild.table) + " " + child.grand.length + "건과 함께 삭제됩니다.";
        }
        card.appendChild(el("div", { class: "ppaf-childcard-note" }, [note]));
      }

      return card;
    }

    function buildGrandchildSection(childDef, child) {
      var gdef = childDef.grandchild;
      var wrap = el("div", { class: "ppaf-grandwrap" });
      var activeCount = child.grand.filter(function (g) { return !g.deleted; }).length;
      var head = el("div", { class: "ppaf-group-headrow" }, [
        el("div", { class: "ppaf-grouptitle small" }, [groupLabel(gdef.table) + " (" + activeCount + "건)"])
      ]);
      var addBtn = el("button", { class: "ppaf-btn", type: "button" }, ["+ " + groupLabel(gdef.table) + " 추가"]);
      addBtn.addEventListener("click", function () {
        child.grand.push({ table: gdef.table, fk: gdef.fk, record: {}, existing: false, deleted: false });
        groupDirty = true;
        renderGroup().catch(console.error);
      });
      head.appendChild(addBtn);
      wrap.appendChild(head);

      child.grand.forEach(function (g) {
        var gcard = el("div", { class: "ppaf-grandcard" + (g.deleted ? " deleted" : "") });
        var gbadge = el("span", { class: "ppaf-badge " + (g.existing ? "existing" : "new") }, [g.existing ? "기존" : "신규"]);
        var gToggle;
        if (g.existing) {
          gToggle = el("button", { class: "ppaf-btn danger", type: "button" }, [g.deleted ? "삭제 취소" : "삭제 표시"]);
          gToggle.addEventListener("click", function () { g.deleted = !g.deleted; groupDirty = true; renderGroup().catch(console.error); });
        } else {
          gToggle = el("button", { class: "ppaf-btn danger", type: "button" }, ["제거"]);
          gToggle.addEventListener("click", function () {
            var i = child.grand.indexOf(g);
            if (i !== -1) child.grand.splice(i, 1);
            groupDirty = true;
            renderGroup().catch(console.error);
          });
        }
        gcard.appendChild(el("div", { class: "ppaf-childcard-head" }, [gbadge, gToggle]));
        if (!g.deleted) {
          gcard.appendChild(buildFieldsGrid(gdef.table, g.record, { exclude: [gdef.fk], alreadyLoaded: g.existing }));
        } else {
          gcard.appendChild(el("div", { class: "ppaf-childcard-note" }, ["이 항목은 저장 시 삭제됩니다."]));
        }
        wrap.appendChild(gcard);
      });

      return wrap;
    }

    function buildChildrenSection(def) {
      var childDef = def.child;
      var section = el("div", { class: "ppaf-group-children" });
      var activeCount = groupState.children.filter(function (c) { return !c.deleted; }).length;
      var head = el("div", { class: "ppaf-group-headrow" }, [
        el("div", { class: "ppaf-grouptitle" }, [groupLabel(childDef.table) + " (" + activeCount + "건)"])
      ]);
      var addBtn = el("button", { class: "ppaf-btn", type: "button" }, ["+ " + groupLabel(childDef.table) + " 추가"]);
      addBtn.addEventListener("click", function () {
        groupState.children.push({ table: childDef.table, fk: childDef.fk, record: {}, existing: false, deleted: false, grand: [] });
        groupDirty = true;
        renderGroup().catch(console.error);
      });
      head.appendChild(addBtn);
      section.appendChild(head);

      groupState.children.forEach(function (child) {
        section.appendChild(buildChildCard(def, child));
      });

      return section;
    }

    function updateGroupDeleteButtonState() {
      if (appMode === "single") return;
      var loaded = !!(groupState && groupState.master.existing);
      deleteBtn.disabled = !loaded;
      deleteBtn.title = loaded ? "" : "먼저 검색으로 기존 데이터를 불러오세요.";
    }

    async function renderGroup(autofocusMaster) {
      var def = GROUP_DEFS[groupState.kind];
      var childDef = def.child;
      var gdef = childDef.grandchild;

      await getTableOptions(def.master, false);
      await getTableOptions(childDef.table, false);
      if (gdef) await getTableOptions(gdef.table, false);
      await getTableRecords(def.master, false);

      groupWrap.innerHTML = "";
      groupWrap.appendChild(
        el("div", { class: "ppaf-help" }, [
          "마스터(" + groupLabel(def.master) + ")와 하위 항목을 함께 검증한 뒤, 하나라도 실패하면 아무것도 반영하지 않고 " +
          "전부 통과해야 한 번에 엑셀에 저장됩니다."
        ])
      );
      var masterToolbar = buildGroupMasterToolbar(def);
      groupWrap.appendChild(masterToolbar);
      groupWrap.appendChild(el("div", { class: "ppaf-grouptitle" }, [groupLabel(def.master)]));
      groupWrap.appendChild(
        buildFieldsGrid(def.master, groupState.master.record, { alreadyLoaded: groupState.master.existing })
      );
      groupWrap.appendChild(buildChildrenSection(def));

      updateGroupDeleteButtonState();

      if (autofocusMaster) {
        var q = masterToolbar.querySelector(".ppaf-picker-q");
        if (q) setTimeout(function () { q.focus(); }, 30);
      }
    }

    // 반환값: 실제로 전환했으면 true, 저장 안 한 변경 확인창에서 사용자가
    // 취소해 전환하지 않았으면 false - 외부(표 탭)에서 호출할 때 이어서
    // 레코드를 불러와도 되는지 판단하는 데 씁니다.
    async function switchMode(mode) {
      if (mode !== appMode && !confirmDiscardIfDirty("저장하지 않은 변경 내용이 있습니다. 화면을 바꾸면 사라집니다. 계속할까요?")) {
        return false;
      }
      formDirty = false;
      groupDirty = false;
      appMode = mode;
      modeBtnSingle.classList.toggle("on", mode === "single");
      modeBtnA.classList.toggle("on", mode === "groupA");
      modeBtnB.classList.toggle("on", mode === "groupB");

      if (mode === "single") {
        singleWrap.style.display = "";
        groupWrap.style.display = "none";
        saveBtn.textContent = "엑셀에 저장";
        deleteBtn.textContent = "삭제";
        await renderMode(true, true);
      } else {
        singleWrap.style.display = "none";
        groupWrap.style.display = "";
        saveBtn.textContent = "일괄 저장";
        deleteBtn.textContent = "그룹 전체 삭제";
        groupState = newGroupState(mode);
        await renderGroup(true);
      }
      return true;
    }

    async function saveGroup() {
      var def = GROUP_DEFS[groupState.kind];
      var childDef = def.child;
      var gdef = childDef.grandchild;
      var masterSchema = SCHEMA_BY_KEY[def.master];
      var childSchema = SCHEMA_BY_KEY[childDef.table];
      var masterPk = masterSchema.pk;

      var errors = [];
      var masterPkVal = (groupState.master.record[masterPk] || "").toString().trim();
      if (!masterPkVal) errors.push(groupLabel(def.master) + "의 " + masterPk + "를 입력하세요.");
      validateRecordRequired(def.master, groupState.master.record).forEach(function (m) {
        errors.push(groupLabel(def.master) + ": " + m);
      });

      var activeChildren = groupState.children.filter(function (c) { return !c.deleted; });
      activeChildren.forEach(function (c, i) {
        c.record[childDef.fk] = masterPkVal;
        validateRecordRequired(childDef.table, c.record).forEach(function (m) {
          errors.push(groupLabel(childDef.table) + " #" + (i + 1) + ": " + m);
        });
        if (gdef) {
          var childPkVal = (c.record[childSchema.pk] || "").toString().trim();
          var activeGrand = (c.grand || []).filter(function (g) { return !g.deleted; });
          activeGrand.forEach(function (g, gi) {
            g.record[gdef.fk] = childPkVal;
            validateRecordRequired(gdef.table, g.record).forEach(function (m) {
              errors.push(groupLabel(gdef.table) + " #" + (i + 1) + "-" + (gi + 1) + ": " + m);
            });
          });
        }
      });

      if (errors.length) {
        showToast("입력값을 확인해주세요: " + errors[0] + (errors.length > 1 ? " 외 " + (errors.length - 1) + "건" : ""), "error");
        return;
      }

      var operations = [{ table: def.master, action: "save", record: groupState.master.record }];
      activeChildren.forEach(function (c) {
        operations.push({ table: childDef.table, action: "save", record: c.record });
        if (gdef) {
          (c.grand || []).filter(function (g) { return !g.deleted; }).forEach(function (g) {
            operations.push({ table: gdef.table, action: "save", record: g.record });
          });
        }
      });
      groupState.children.filter(function (c) { return c.deleted; }).forEach(function (c) {
        if (gdef) {
          (c.grand || []).filter(function (g) { return g.existing; }).forEach(function (g) {
            operations.push({ table: gdef.table, action: "delete", pk: g.record[SCHEMA_BY_KEY[gdef.table].pk] });
          });
        }
        operations.push({ table: childDef.table, action: "delete", pk: c.record[childSchema.pk] });
      });

      try {
        var data = await withBusy(saveBtn, "일괄 저장 중...", function () {
          return apiPost("/api/batch", { operations: operations, actor: getActorName() });
        });
        showToast(data.message || "일괄 저장 완료", "success");
        groupDirty = false;
        delete recordCache[def.master];
        delete optionCache[def.master];
        delete recordCache[childDef.table];
        delete optionCache[childDef.table];
        if (gdef) { delete recordCache[gdef.table]; delete optionCache[gdef.table]; }
        rememberDashboardTab();
        setTimeout(function () { location.reload(); }, 700);
      } catch (e) {
        showToast("일괄 저장 실패: " + (e.message || e), "error");
      }
    }

    async function performGroupDelete() {
      if (!groupState || !groupState.master.existing) return;
      var def = GROUP_DEFS[groupState.kind];
      var childDef = def.child;
      var gdef = childDef.grandchild;
      var masterSchema = SCHEMA_BY_KEY[def.master];
      var childSchema = SCHEMA_BY_KEY[childDef.table];
      var masterPkVal = groupState.master.record[masterSchema.pk];

      var confirmed = window.confirm(
        groupLabel(def.master) + "(" + masterPkVal + ")와(과) 연결된 하위 데이터를 전부 삭제합니다. " +
        "되돌릴 수 없습니다. 계속할까요?"
      );
      if (!confirmed) return;

      var operations = [];
      groupState.children.forEach(function (c) {
        if (!c.existing) return;
        if (gdef) {
          (c.grand || []).forEach(function (g) {
            if (g.existing) operations.push({ table: gdef.table, action: "delete", pk: g.record[SCHEMA_BY_KEY[gdef.table].pk] });
          });
        }
        operations.push({ table: childDef.table, action: "delete", pk: c.record[childSchema.pk] });
      });
      operations.push({ table: def.master, action: "delete", pk: masterPkVal });

      try {
        var data = await withBusy(deleteBtn, "삭제 중...", function () {
          return apiPost("/api/batch", { operations: operations, actor: getActorName() });
        });
        showToast(data.message || "그룹 삭제 완료", "success");
        groupDirty = false;
        delete recordCache[def.master];
        delete optionCache[def.master];
        delete recordCache[childDef.table];
        delete optionCache[childDef.table];
        if (gdef) { delete recordCache[gdef.table]; delete optionCache[gdef.table]; }
        groupState = newGroupState(def.key);
        rememberDashboardTab();
        setTimeout(function () { location.reload(); }, 700);
      } catch (e) {
        showToast("그룹 삭제 실패: " + (e.message || e), "error");
      }
    }

    async function renderMode(forceReload, autofocusPicker) {
      var tableName = currentTable();
      fields.innerHTML = "";
      formState.loadedPk = null;
      formDirty = false;
      help.textContent = "저장하면 지금 열려 있는(또는 열려 있지 않으면 새로 열리는) 엑셀 파일에 바로 반영됩니다.";
      await prepareOptionsForTable(tableName, !!forceReload);
      await getTableRecords(tableName, !!forceReload);
      renderToolbar(tableName, autofocusPicker);
      renderFields(tableName);
      updateDeleteButtonState();
    }

    async function ensureSchema() {
      if (SCHEMA) return;
      var data = await apiGet("/api/schema");
      SCHEMA = data.schema || [];
      SCHEMA.forEach(function (t) { SCHEMA_BY_KEY[t.key] = t; });
    }

    async function openModal() {
      backdrop.classList.add("show");
      modal.classList.add("show");
      try {
        await ensureSchema();
        if (!modeSel.options.length) {
          Object.keys(SCHEMA_BY_KEY).forEach(function (key) {
            modeSel.appendChild(el("option", { value: key }, [TABLE_META[key] ? TABLE_META[key].label : key]));
          });
        }
        await switchMode("single");
      } catch (e) {
        showToast("화면 로딩 실패: " + (e.message || e), "error");
      }
    }

    function closeModal() {
      if (!confirmDiscardIfDirty()) return;
      formDirty = false;
      groupDirty = false;
      backdrop.classList.remove("show");
      modal.classList.remove("show");
    }

    // ---------------------------------------------------------------------
    // 버튼 처리 중 로딩 표시 (스피너 + 비활성화)
    // ---------------------------------------------------------------------
    async function withBusy(btn, busyText, fn) {
      var original = btn.textContent;
      var spinner = el("span", { class: "ppaf-spinner" });
      btn.disabled = true;
      btn.textContent = "";
      btn.appendChild(spinner);
      btn.appendChild(document.createTextNode(" " + busyText));
      try {
        return await fn();
      } finally {
        btn.disabled = false;
        btn.textContent = original;
      }
    }

    async function rebuildDashboard() {
      try {
        await withBusy(rebuildBtn, "새로고침 중...", function () {
          return apiPost("/api/rebuild", {});
        });
        rememberDashboardTab();
        location.reload();
      } catch (e) {
        showToast("새로고침 실패: " + (e.message || e), "error");
      }
    }

    // [변경] 탭의 "기준점 대비 누적 변경"은 이 버튼을 누르기 전까지 계속
    // 쌓여서 보입니다 - 리셋하면 지금 시점이 새 기준점이 됩니다. 여러 생성에
    // 걸쳐 남는 "전체 변경 이력"은 이 리셋과 무관하게 계속 보존됩니다.
    async function resetChangeBaseline() {
      var confirmed = window.confirm(
        "지금 시점을 [변경] 탭의 새 비교 기준점으로 리셋합니다. 그동안 쌓여있던 " +
        "\"기준점 대비 누적 변경\" 표시가 초기화됩니다(전체 변경 이력은 그대로 남습니다). " +
        "계속할까요?"
      );
      if (!confirmed) return;
      try {
        var data = await withBusy(resetBaselineBtn, "리셋 중...", function () {
          return apiPost("/api/reset_snapshot", {});
        });
        showToast(data.message || "리셋 완료", "success");
        rememberDashboardTab();
        setTimeout(function () { location.reload(); }, 700);
      } catch (e) {
        showToast("리셋 실패: " + (e.message || e), "error");
      }
    }

    async function saveCurrentTable() {
      var tableName = currentTable();
      var record = collectRecord();
      var schema = SCHEMA_BY_KEY[tableName];
      var pkName = schema && schema.pk;

      if (!validateAllFields(tableName)) {
        showToast("입력값을 확인해주세요.", "error");
        return;
      }
      if (!pkName || !record[pkName]) {
        showToast((pkName || "PK") + "는 필수입니다.", "error");
        return;
      }

      try {
        var data = await withBusy(saveBtn, "엑셀에 저장 중...", function () {
          return apiPost("/api/save", { table: tableName, record: record, actor: getActorName() });
        });
        showToast(data.message || "저장 완료", "success");

        delete optionCache[tableName];
        delete recordCache[tableName];
        await renderMode(true);
        fillRecord(record);
        formState.loadedPk = record[pkName] || null;
        formDirty = false;
        updateDeleteButtonState();

        rememberDashboardTab();
        setTimeout(function () { location.reload(); }, 700);
      } catch (e) {
        showToast("저장 실패: " + (e.message || e), "error");
      }
    }

    // ---------------------------------------------------------------------
    // 삭제 - 2차 확인 모달 (참조 무결성 경고 포함)
    // ---------------------------------------------------------------------
    async function openDeleteConfirm() {
      var tableName = currentTable();
      var schema = SCHEMA_BY_KEY[tableName];
      var pkValue = formState.loadedPk;
      if (!pkValue || !schema) return;

      confirmBody.innerHTML = "";
      confirmBody.appendChild(el("div", { class: "ppaf-confirm-loading" }, ["연관 데이터 확인 중..."]));
      confirmBackdrop.classList.add("show");
      confirmBox.classList.add("show");
      confirmDeleteBtn.disabled = true;

      var references = [];
      try {
        var data = await apiGet(
          "/api/references?table=" + encodeURIComponent(tableName) + "&pk=" + encodeURIComponent(pkValue)
        );
        references = data.references || [];
      } catch (e) {
        confirmBody.innerHTML = "";
        confirmBody.appendChild(el("div", { class: "ppaf-confirm-error" }, ["연관 데이터 확인 실패: " + (e.message || e)]));
        return;
      }

      confirmBody.innerHTML = "";
      confirmBody.appendChild(
        el("div", { class: "ppaf-confirm-target" }, [
          el("span", { class: "ppaf-confirm-table" }, [TABLE_META[tableName] ? TABLE_META[tableName].label : tableName]),
          " · " + schema.pk + " = ",
          el("strong", {}, [pkValue])
        ])
      );

      if (references.length > 0) {
        var list = el("ul", { class: "ppaf-conflict-list" });
        references.forEach(function (r) {
          list.appendChild(
            el("li", {}, [(TABLE_META[r.table] ? TABLE_META[r.table].label : r.table) + " " + r.count + "건 (" + r.fk_col + ")"])
          );
        });
        confirmBody.appendChild(
          el("div", { class: "ppaf-confirm-warn" }, [
            "다른 표에서 이 데이터를 참조하고 있어 삭제할 수 없습니다:"
          ])
        );
        confirmBody.appendChild(list);
        confirmBody.appendChild(
          el("div", { class: "ppaf-confirm-note" }, ["참조하는 데이터를 먼저 정리한 뒤 다시 시도해주세요."])
        );
        confirmDeleteBtn.disabled = true;
      } else {
        confirmBody.appendChild(
          el("div", { class: "ppaf-confirm-warn" }, ["이 작업은 되돌릴 수 없습니다. 정말 삭제하시겠습니까?"])
        );
        confirmDeleteBtn.disabled = false;
      }
    }

    function closeDeleteConfirm() {
      confirmBackdrop.classList.remove("show");
      confirmBox.classList.remove("show");
    }

    async function performDelete() {
      var tableName = currentTable();
      var schema = SCHEMA_BY_KEY[tableName];
      var pkValue = formState.loadedPk;
      if (!pkValue || !schema) return;

      try {
        var data = await withBusy(confirmDeleteBtn, "삭제 중...", function () {
          return apiPost("/api/delete", { table: tableName, pk: pkValue, actor: getActorName() });
        });
        showToast(data.message || "삭제 완료", "success");
        closeDeleteConfirm();
        clearRecord();
        delete optionCache[tableName];
        delete recordCache[tableName];
        await renderMode(true);
        rememberDashboardTab();
        setTimeout(function () { location.reload(); }, 700);
      } catch (e) {
        showToast("삭제 실패: " + (e.message || e), "error");
        closeDeleteConfirm();
      }
    }

    // ---------------------------------------------------------------------
    // 마크업/스타일
    // ---------------------------------------------------------------------
    var style = el("style", {
      html:
        ":root{--ppaf-teal:#0B8577;--ppaf-teal-d:#086b60;--ppaf-ink:#182422;--ppaf-sub:#5b6b65;--ppaf-line:#dfe8e4;--ppaf-bg:#ffffff;--ppaf-danger:#d92d20;--ppaf-danger-bg:#fdeceb;}" +
        "@media(prefers-color-scheme:dark){:root{--ppaf-ink:#e9f3f0;--ppaf-sub:#a7b6b1;--ppaf-line:#2b3937;--ppaf-bg:#182422;}}" +
        ".ppaf-open{position:fixed;right:24px;bottom:24px;z-index:2147483000;border:none;border-radius:999px;padding:13px 18px;background:var(--ppaf-teal);color:#fff;font-weight:800;font-size:14px;cursor:pointer;box-shadow:0 10px 28px rgba(11,133,119,.35);transition:transform .15s ease,box-shadow .15s ease}" +
        ".ppaf-open:hover{transform:translateY(-1px);box-shadow:0 14px 32px rgba(11,133,119,.4)}" +
        ".ppaf-backdrop,.ppaf-confirm-backdrop{position:fixed;inset:0;background:rgba(10,20,18,.45);z-index:2147482998;display:none;backdrop-filter:blur(1px)}" +
        ".ppaf-backdrop.show,.ppaf-confirm-backdrop.show{display:block}" +
        ".ppaf-modal{position:fixed;right:24px;bottom:80px;width:min(1100px,calc(100vw - 32px));max-height:84vh;overflow:auto;background:var(--ppaf-bg);color:var(--ppaf-ink);border:1px solid var(--ppaf-line);border-radius:18px;box-shadow:0 24px 60px rgba(0,0,0,.28);z-index:2147482999;display:none;opacity:0;transform:translateY(8px);transition:opacity .18s ease,transform .18s ease}" +
        ".ppaf-modal.show{display:block;opacity:1;transform:translateY(0)}" +
        ".ppaf-head{padding:18px 20px;border-bottom:1px solid var(--ppaf-line);display:flex;gap:12px;justify-content:space-between;align-items:center;position:sticky;top:0;background:var(--ppaf-bg);border-radius:18px 18px 0 0;flex-wrap:wrap}" +
        ".ppaf-actor-wrap{display:flex;align-items:center;gap:5px;flex:0 1 200px}" +
        ".ppaf-actor-label{font-size:13px}" +
        ".ppaf-actor-input{padding:6px 9px;font-size:12.5px;min-width:0}" +
        ".ppaf-title{flex:1 1 auto;font-size:16px;font-weight:800}" +
        ".ppaf-body{padding:18px 20px 22px}" +
        ".ppaf-help{font-size:12px;color:var(--ppaf-sub);margin:10px 0 16px;line-height:1.5}" +
        ".ppaf-fields{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px}" +
        ".ppaf-row{display:flex;flex-direction:column;gap:5px}" +
        ".ppaf-label{font-size:12px;font-weight:700;color:var(--ppaf-ink)}" +
        ".ppaf-input,.ppaf-textarea{font-size:13.5px;padding:10px 11px;border:1.5px solid var(--ppaf-line);border-radius:9px;background:var(--ppaf-bg);color:var(--ppaf-ink);transition:border-color .12s ease,box-shadow .12s ease}" +
        ".ppaf-input:focus,.ppaf-textarea:focus{outline:none;border-color:var(--ppaf-teal);box-shadow:0 0 0 3px rgba(11,133,119,.15)}" +
        ".ppaf-textarea{min-height:88px;resize:vertical}" +
        ".ppaf-row.invalid .ppaf-input,.ppaf-row.invalid .ppaf-textarea{border-color:var(--ppaf-danger);box-shadow:0 0 0 3px rgba(217,45,32,.12)}" +
        ".ppaf-fieldnote{font-size:11px;line-height:1.4}" +
        ".ppaf-fieldnote-error{color:var(--ppaf-danger)}" +
        ".ppaf-fieldnote-hint{color:#9a6b00}" +
        ".ppaf-foot{padding:16px 20px;border-top:1px solid var(--ppaf-line);display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap;position:sticky;bottom:0;background:var(--ppaf-bg);border-radius:0 0 18px 18px}" +
        ".ppaf-btn{border:1.5px solid var(--ppaf-line);background:var(--ppaf-bg);color:var(--ppaf-ink);border-radius:9px;padding:10px 14px;cursor:pointer;font-weight:700;font-size:13px;display:inline-flex;align-items:center;gap:6px;transition:background .12s ease,border-color .12s ease}" +
        ".ppaf-btn:hover:not(:disabled){border-color:var(--ppaf-teal)}" +
        ".ppaf-btn:disabled{opacity:.45;cursor:not-allowed}" +
        ".ppaf-btn.primary{background:var(--ppaf-teal);border-color:var(--ppaf-teal);color:#fff}" +
        ".ppaf-btn.primary:hover:not(:disabled){background:var(--ppaf-teal-d)}" +
        ".ppaf-btn.danger{background:var(--ppaf-bg);border-color:var(--ppaf-danger);color:var(--ppaf-danger)}" +
        ".ppaf-btn.danger:hover:not(:disabled){background:var(--ppaf-danger-bg)}" +
        ".ppaf-required::after{content:' *';color:var(--ppaf-danger);font-weight:800}" +
        ".ppaf-formula24-display{font-size:11px;color:var(--ppaf-sub);margin-top:1px}" +
        ".ppaf-toolbar-wrap{display:flex;flex-direction:column;gap:6px;margin:4px 0 18px;grid-column:1 / -1}" +
        ".ppaf-picker-label{font-size:11.5px;font-weight:700;color:var(--ppaf-sub);margin-top:4px}" +
        ".ppaf-picker{border:1.5px solid var(--ppaf-line);border-radius:10px;padding:10px;background:rgba(11,133,119,.03)}" +
        ".ppaf-picker-controls{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px}" +
        ".ppaf-picker-col{flex:0 0 auto;min-width:120px}" +
        ".ppaf-picker-q{flex:1 1 220px}" +
        ".ppaf-picker-results{max-height:220px;overflow:auto;display:flex;flex-direction:column;gap:4px}" +
        ".ppaf-picker-item{text-align:left;border:1px solid transparent;background:var(--ppaf-bg);color:var(--ppaf-ink);border-radius:7px;padding:8px 10px;font-size:12.5px;cursor:pointer;transition:background .12s ease,border-color .12s ease}" +
        ".ppaf-picker-item:hover{background:rgba(11,133,119,.1);border-color:var(--ppaf-teal)}" +
        ".ppaf-picker-hint,.ppaf-picker-more{color:var(--ppaf-sub);font-size:12px;padding:6px 2px}" +
        ".ppaf-loaded-indicator{display:flex;align-items:center;gap:8px;background:rgba(11,133,119,.08);border:1px solid rgba(11,133,119,.25);border-radius:999px;padding:6px 8px 6px 12px;font-size:12.5px;font-weight:600;color:var(--ppaf-teal-d)}" +
        ".ppaf-loaded-dot{width:7px;height:7px;border-radius:50%;background:var(--ppaf-teal);flex:0 0 auto}" +
        ".ppaf-loaded-label{flex:1 1 auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}" +
        ".ppaf-chip-x{border:none;background:transparent;color:var(--ppaf-teal-d);cursor:pointer;font-size:13px;line-height:1;padding:3px 5px;border-radius:50%;flex:0 0 auto}" +
        ".ppaf-chip-x:hover{background:rgba(11,133,119,.18)}" +
        ".ppaf-modeswitch{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}" +
        ".ppaf-modebtn{border:1.5px solid var(--ppaf-line);background:var(--ppaf-bg);color:var(--ppaf-sub);border-radius:999px;padding:8px 14px;font-size:12.5px;font-weight:700;cursor:pointer;transition:all .12s ease}" +
        ".ppaf-modebtn.on{background:var(--ppaf-teal);border-color:var(--ppaf-teal);color:#fff}" +
        ".ppaf-modebtn:hover:not(.on){border-color:var(--ppaf-teal);color:var(--ppaf-ink)}" +
        ".ppaf-grouptitle{font-size:13.5px;font-weight:800;margin:6px 0 10px}" +
        ".ppaf-grouptitle.small{font-size:12px;margin:0}" +
        ".ppaf-group-headrow{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px}" +
        ".ppaf-group-children{margin-top:18px;padding-top:14px;border-top:1px dashed var(--ppaf-line)}" +
        ".ppaf-childcard{border:1.5px solid var(--ppaf-line);border-radius:12px;padding:12px 14px;margin-bottom:12px;background:rgba(11,133,119,.02)}" +
        ".ppaf-childcard.deleted{background:var(--ppaf-danger-bg);border-color:var(--ppaf-danger)}" +
        ".ppaf-childcard-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}" +
        ".ppaf-childcard-note{color:var(--ppaf-danger);font-size:12.5px;font-weight:600}" +
        ".ppaf-badge{font-size:10.5px;font-weight:800;padding:3px 8px;border-radius:999px}" +
        ".ppaf-badge.existing{background:rgba(11,133,119,.14);color:var(--ppaf-teal-d)}" +
        ".ppaf-badge.new{background:rgba(154,107,0,.14);color:#9a6b00}" +
        ".ppaf-grandwrap{margin-top:10px;padding:10px;border:1px dashed var(--ppaf-line);border-radius:10px}" +
        ".ppaf-grandcard{border:1px solid var(--ppaf-line);border-radius:9px;padding:10px 12px;margin-bottom:8px;background:var(--ppaf-bg)}" +
        ".ppaf-grandcard.deleted{background:var(--ppaf-danger-bg);border-color:var(--ppaf-danger)}" +
        ".ppaf-spinner{display:inline-block;width:13px;height:13px;border-radius:50%;border:2px solid rgba(255,255,255,.5);border-top-color:#fff;animation:ppaf-spin .7s linear infinite;vertical-align:-2px}" +
        ".ppaf-btn:not(.primary):not(.danger) .ppaf-spinner{border-color:rgba(11,133,119,.3);border-top-color:var(--ppaf-teal)}" +
        "@keyframes ppaf-spin{to{transform:rotate(360deg)}}" +
        ".ppaf-toaststack{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);z-index:2147483647;display:flex;flex-direction:column;gap:8px;align-items:center;pointer-events:none}" +
        ".ppaf-toast{pointer-events:auto;max-width:min(420px,86vw);background:var(--ppaf-ink);color:#fff;padding:11px 16px;border-radius:10px;font-size:13px;font-weight:600;box-shadow:0 10px 26px rgba(0,0,0,.28);opacity:0;transform:translateY(8px);transition:opacity .2s ease,transform .2s ease}" +
        ".ppaf-toast.show{opacity:1;transform:translateY(0)}" +
        ".ppaf-toast-success{background:#0B8577}" +
        ".ppaf-toast-error{background:#c0362b}" +
        ".ppaf-confirm-backdrop{z-index:2147483001}" +
        ".ppaf-confirm{position:fixed;left:50%;top:50%;transform:translate(-50%,-46%);width:min(420px,90vw);background:var(--ppaf-bg);color:var(--ppaf-ink);border-radius:16px;border:1px solid var(--ppaf-line);box-shadow:0 24px 60px rgba(0,0,0,.32);z-index:2147483002;display:none;opacity:0;transition:opacity .16s ease,transform .16s ease}" +
        ".ppaf-confirm.show{display:block;opacity:1;transform:translate(-50%,-50%)}" +
        ".ppaf-confirm-head{display:flex;align-items:center;gap:10px;padding:18px 20px 10px;font-weight:800;font-size:15px}" +
        ".ppaf-confirm-icon{font-size:20px}" +
        ".ppaf-confirm-body{padding:0 20px 6px;font-size:13px;line-height:1.6}" +
        ".ppaf-confirm-target{background:rgba(11,133,119,.08);border-radius:8px;padding:8px 10px;margin-bottom:10px;font-size:12.5px}" +
        ".ppaf-confirm-table{font-weight:700}" +
        ".ppaf-confirm-warn{font-weight:700;margin-bottom:6px}" +
        ".ppaf-conflict-list{margin:0 0 8px;padding-left:18px;color:var(--ppaf-danger)}" +
        ".ppaf-confirm-note{color:var(--ppaf-sub);font-size:12px}" +
        ".ppaf-confirm-error{color:var(--ppaf-danger)}" +
        ".ppaf-confirm-loading{color:var(--ppaf-sub)}" +
        ".ppaf-confirm-foot{display:flex;justify-content:flex-end;gap:8px;padding:16px 20px 20px}" +
        "@media(max-width:640px){.ppaf-modal{right:12px;bottom:72px;width:calc(100vw - 24px)}.ppaf-fields{grid-template-columns:1fr}.ppaf-picker-controls{flex-direction:column}.ppaf-open{right:16px;bottom:16px;padding:12px 16px;font-size:13px}}" +
        ".ppaf-matchtrigger{background:rgba(11,133,119,.08);border-color:var(--ppaf-teal);color:var(--ppaf-teal-d)}" +
        ".ppaf-matchtrigger:hover:not(:disabled){background:rgba(11,133,119,.16)}" +
        ".ppaf-matchwrap{grid-column:1 / -1;margin:10px 0 16px;border:1.5px solid var(--ppaf-teal);border-radius:12px;padding:14px;background:rgba(11,133,119,.04)}" +
        ".ppaf-matchgrid{display:grid;grid-template-columns:1fr 1fr;gap:14px}" +
        ".ppaf-matchcol-title{font-size:12px;font-weight:800;color:var(--ppaf-teal-d);margin-bottom:6px}" +
        ".ppaf-matchsearch{width:100%;box-sizing:border-box;margin-bottom:8px}" +
        ".ppaf-matchlist{max-height:260px;overflow:auto;display:flex;flex-direction:column;gap:6px;padding-right:2px}" +
        ".ppaf-matchcard{border:1.5px solid var(--ppaf-line);border-radius:9px;padding:8px 10px;background:var(--ppaf-bg);cursor:pointer;transition:border-color .12s ease,background .12s ease}" +
        ".ppaf-matchcard:hover{border-color:var(--ppaf-teal)}" +
        ".ppaf-matchcard.selected{border-color:var(--ppaf-teal);background:rgba(11,133,119,.12);box-shadow:0 0 0 2px rgba(11,133,119,.18)}" +
        ".ppaf-matchcard-head{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:13px}" +
        ".ppaf-matchcard-line{font-size:12px;color:var(--ppaf-ink);margin-top:2px}" +
        ".ppaf-matchcard-sub{font-size:11px;color:var(--ppaf-sub);margin-top:1px}" +
        ".ppaf-matchpill{font-size:10.5px;font-weight:700;color:var(--ppaf-sub);background:rgba(91,107,101,.12);border-radius:999px;padding:2px 8px;white-space:nowrap}" +
        ".ppaf-matchpill.on{color:var(--ppaf-teal-d);background:rgba(11,133,119,.14)}" +
        ".ppaf-matchfoot{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:12px;padding-top:12px;border-top:1px dashed var(--ppaf-line);flex-wrap:wrap}" +
        ".ppaf-matchpreview{font-size:12.5px;font-weight:700;color:var(--ppaf-sub)}" +
        ".ppaf-matchpreview.ok{color:var(--ppaf-teal-d)}" +
        ".ppaf-matchpreview.warn{color:var(--ppaf-danger)}" +
        ".ppaf-matchfoot-btns{display:flex;gap:8px}" +
        "@media(max-width:640px){.ppaf-matchgrid{grid-template-columns:1fr}}"
    });

    var openBtn = el("button", { class: "ppaf-open", type: "button" }, ["✎ 간편 입력/저장"]);
    var backdrop = el("div", { class: "ppaf-backdrop" });
    var modal = el("div", { class: "ppaf-modal" });

    var closeBtn = el("button", { class: "ppaf-btn", type: "button" }, ["닫기"]);
    var clearBtn = el("button", { class: "ppaf-btn", type: "button" }, ["초기화"]);
    var rebuildBtn = el("button", {
      class: "ppaf-btn", type: "button",
      title: "다른 사람이 방금 저장한 내용이 안 보일 때 눌러주세요 - 그래도 안 보이면 엑셀 파일을 직접 닫았다 다시 열어야 확실합니다."
    }, ["대시보드 새로고침"]);
    var resetBaselineBtn = el("button", { class: "ppaf-btn", type: "button", title: "[변경] 탭의 비교 기준점을 지금 시점으로 리셋합니다" }, ["변경 비교 기준 리셋"]);
    var deleteBtn = el("button", { class: "ppaf-btn danger", type: "button", disabled: "disabled" }, ["삭제"]);
    var saveBtn = el("button", { class: "ppaf-btn primary", type: "button" }, ["엑셀에 저장"]);

    var modeSel = el("select", { class: "ppaf-input" });

    var help = el("div", { class: "ppaf-help" });
    var bodyWrap = el("div", { class: "ppaf-body" });
    var fields = el("div", { class: "ppaf-fields" });

    var singleWrap = el("div", { class: "ppaf-singlewrap" });
    singleWrap.appendChild(modeSel);
    singleWrap.appendChild(help);
    singleWrap.appendChild(fields);

    var groupWrap = el("div", { class: "ppaf-groupwrap" });
    groupWrap.style.display = "none";

    var modeBtnSingle = el("button", { class: "ppaf-modebtn on", type: "button" }, ["개별입력"]);
    var modeBtnA = el("button", { class: "ppaf-modebtn", type: "button" }, [GROUP_DEFS.groupA.label]);
    var modeBtnB = el("button", { class: "ppaf-modebtn", type: "button" }, [GROUP_DEFS.groupB.label]);
    var modeSwitchWrap = el("div", { class: "ppaf-modeswitch" }, [modeBtnSingle, modeBtnA, modeBtnB]);

    var actorInput = el("input", {
      class: "ppaf-input ppaf-actor-input",
      type: "text",
      placeholder: "표시 이름(선택)",
      title: "저장/삭제 이력에 남길 내 이름 - 비워두면 이 컴퓨터의 Windows 로그인 계정이 자동으로 쓰입니다."
    });
    actorInput.value = getActorName();
    actorInput.addEventListener("change", function () { setActorName(actorInput.value); });
    actorInput.addEventListener("blur", function () { setActorName(actorInput.value); });

    modal.appendChild(el("div", { class: "ppaf-head" }, [
      el("div", { class: "ppaf-title" }, ["간편 입력/저장 (엑셀에 바로 반영)"]),
      el("div", { class: "ppaf-actor-wrap" }, [el("span", { class: "ppaf-actor-label" }, ["👤"]), actorInput]),
      closeBtn
    ]));
    bodyWrap.appendChild(modeSwitchWrap);
    bodyWrap.appendChild(singleWrap);
    bodyWrap.appendChild(groupWrap);
    modal.appendChild(bodyWrap);
    modal.appendChild(el("div", { class: "ppaf-foot" }, [clearBtn, deleteBtn, resetBaselineBtn, rebuildBtn, saveBtn]));

    // 삭제 2차 확인 모달
    var confirmBackdrop = el("div", { class: "ppaf-confirm-backdrop" });
    var confirmBox = el("div", { class: "ppaf-confirm" });
    var confirmBody = el("div", { class: "ppaf-confirm-body" });
    var confirmCancelBtn = el("button", { class: "ppaf-btn", type: "button" }, ["취소"]);
    var confirmDeleteBtn = el("button", { class: "ppaf-btn danger", type: "button" }, ["삭제"]);
    confirmBox.appendChild(
      el("div", { class: "ppaf-confirm-head" }, [el("span", { class: "ppaf-confirm-icon" }, ["⚠️"]), "정말 삭제할까요?"])
    );
    confirmBox.appendChild(confirmBody);
    confirmBox.appendChild(el("div", { class: "ppaf-confirm-foot" }, [confirmCancelBtn, confirmDeleteBtn]));

    document.head.appendChild(style);
    document.body.appendChild(backdrop);
    document.body.appendChild(modal);
    document.body.appendChild(confirmBackdrop);
    document.body.appendChild(confirmBox);
    document.body.appendChild(toastStack);
    document.body.appendChild(openBtn);

    openBtn.addEventListener("click", function () { openModal().catch(console.error); });
    closeBtn.addEventListener("click", closeModal);
    backdrop.addEventListener("click", closeModal);
    clearBtn.addEventListener("click", function () {
      if (!confirmDiscardIfDirty()) return;
      if (appMode === "single") {
        clearRecord();
      } else {
        groupDirty = false;
        groupState = newGroupState(appMode);
        renderGroup().catch(console.error);
      }
    });
    rebuildBtn.addEventListener("click", function () { rebuildDashboard().catch(console.error); });
    resetBaselineBtn.addEventListener("click", function () { resetChangeBaseline().catch(console.error); });
    saveBtn.addEventListener("click", function () {
      if (appMode === "single") saveCurrentTable().catch(console.error);
      else saveGroup().catch(console.error);
    });
    deleteBtn.addEventListener("click", function () {
      if (appMode === "single") openDeleteConfirm().catch(console.error);
      else performGroupDelete().catch(console.error);
    });
    confirmCancelBtn.addEventListener("click", closeDeleteConfirm);
    confirmBackdrop.addEventListener("click", closeDeleteConfirm);
    confirmDeleteBtn.addEventListener("click", function () { performDelete().catch(console.error); });
    modeSel.addEventListener("change", function () {
      if (!confirmDiscardIfDirty()) { modeSel.value = currentTable(); return; }
      renderMode(true, true).catch(function (e) { showToast("목록 로딩 실패: " + (e.message || e), "error"); });
    });
    modeBtnSingle.addEventListener("click", function () { switchMode("single").catch(console.error); });
    modeBtnA.addEventListener("click", function () { switchMode("groupA").catch(console.error); });
    modeBtnB.addEventListener("click", function () { switchMode("groupB").catch(console.error); });

    // 실제 데이터 입력칸(.ppaf-row 안)에서만 편집 상태를 추적합니다 - 검색
    // 상자/컬럼 선택 등 툴바 요소는 "편집"으로 치지 않습니다.
    fields.addEventListener("input", function (e) {
      if (e.target && e.target.closest && e.target.closest(".ppaf-row")) formDirty = true;
    });
    groupWrap.addEventListener("input", function (e) {
      if (e.target && e.target.closest && e.target.closest(".ppaf-row")) groupDirty = true;
    });

    // 입력칸에서 Enter → 저장 (검색 상자 제외, textarea는 줄바꿈 유지)
    modal.addEventListener("keydown", function (e) {
      if (e.key !== "Enter") return;
      var t = e.target;
      if (!t || t.tagName === "TEXTAREA") return;
      if (t.classList && (t.classList.contains("ppaf-picker-q") || t.classList.contains("ppaf-picker-col"))) return;
      if (t.tagName === "INPUT" || t.tagName === "SELECT") {
        e.preventDefault();
        saveBtn.click();
      }
    });

    document.addEventListener("keydown", function (e) {
      if (e.key !== "Escape") return;
      if (confirmBox.classList.contains("show")) { closeDeleteConfirm(); return; }
      if (modal.classList.contains("show")) { closeModal(); }
    });

    // ---------------------------------------------------------------------
    // 대시보드 각 표 탭(상세 모달의 "수정"/"삭제", 목록 위의 "+ 추가")에서
    // 이 플로팅 폼을 원격으로 여는 진입점. window.PPA_FORM 이 있는지로 "지금
    // 실시간 입력 서버가 붙어 있는지"를 판단하므로, 정적으로 생성된 HTML을
    // 서버 없이 열었을 때는 이 객체 자체가 없어 대시보드 쪽에서 안내 메시지로
    // 대체합니다.
    // ---------------------------------------------------------------------
    async function externalSelectTable(tableKey) {
      await ensureSchema();
      if (!modeSel.options.length) {
        Object.keys(SCHEMA_BY_KEY).forEach(function (key) {
          modeSel.appendChild(el("option", { value: key }, [TABLE_META[key] ? TABLE_META[key].label : key]));
        });
      }
      if (!SCHEMA_BY_KEY[tableKey]) throw new Error("알 수 없는 표: " + tableKey);

      if (!modal.classList.contains("show")) {
        backdrop.classList.add("show");
        modal.classList.add("show");
      }

      if (appMode !== "single") {
        var switched = await switchMode("single");
        if (!switched) return false;
      } else if (modeSel.value !== tableKey && !confirmDiscardIfDirty()) {
        return false;
      }

      if (modeSel.value !== tableKey) {
        modeSel.value = tableKey;
        await renderMode(true, false);
      }
      return true;
    }

    window.PPA_FORM = {
      edit: function (tableKey, pkValue) {
        return externalSelectTable(tableKey).then(function (ok) {
          if (!ok || !pkValue) return;
          return apiGet("/api/record?table=" + encodeURIComponent(tableKey) + "&pk=" + encodeURIComponent(pkValue))
            .then(function (data) { applyLoadedRecord(tableKey, data.record || {}); });
        }).catch(function (e) { showToast("불러오기 실패: " + (e.message || e), "error"); });
      },
      add: function (tableKey) {
        return externalSelectTable(tableKey).catch(function (e) {
          showToast("화면 전환 실패: " + (e.message || e), "error");
        });
      },
      del: function (tableKey, pkValue) {
        return externalSelectTable(tableKey).then(function (ok) {
          if (!ok || !pkValue) return;
          return apiGet("/api/record?table=" + encodeURIComponent(tableKey) + "&pk=" + encodeURIComponent(pkValue))
            .then(function (data) {
              applyLoadedRecord(tableKey, data.record || {});
              return openDeleteConfirm();
            });
        }).catch(function (e) { showToast("불러오기 실패: " + (e.message || e), "error"); });
      },
      resetBaseline: function () {
        return resetChangeBaseline();
      }
    };

    // 대시보드 탭/창을 진짜로 닫으면 실시간 입력 서버도 같이 종료합니다 -
    // 매번 stop_live_server.bat 을 따로 실행해야 하는 불편함을 없애기
    // 위함입니다. pagehide는 우리가 스스로 부르는 location.reload()에서도
    // 똑같이 발생하므로, rememberDashboardTab()이 남겨둔 표시로 "저장 후
    // 새로고침"과 "진짜 닫기"를 구분합니다(표시가 있으면 지우고 아무것도
    // 안 보냄). event.persisted(bfcache로 페이지가 보존되는 경우, 진짜 닫힘이
    // 아님)일 때도 보내지 않습니다. sendBeacon은 페이지가 사라지는 순간에도
    // 브라우저가 요청 전송을 보장해주는 유일한 방법이라 이 용도에 씁니다.
    window.addEventListener("pagehide", function (event) {
      if (event.persisted) return;
      try {
        if (sessionStorage.getItem("ppa_intentional_reload") === "1") {
          sessionStorage.removeItem("ppa_intentional_reload");
          return;
        }
      } catch (e) { /* 세션스토리지 사용 불가 환경 - 판단 불가하니 그냥 보냄 */ }
      try { navigator.sendBeacon("/api/shutdown"); } catch (e) { /* 무시 */ }
    });

    console.log("[ppa] dashboard_form.js loaded");
  } catch (e) {
    console.error("[ppa] dashboard_form.js error:", e);
  }
})();
