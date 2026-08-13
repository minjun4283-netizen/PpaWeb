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

    // loadedPk: 지금 폼이 "기존 데이터를 불러온" 상태인지(삭제 가능) 아니면
    // "새 입력" 상태인지(삭제 버튼 비활성) 추적합니다.
    var formState = { loadedPk: null };

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

    async function apiGet(url) {
      var res = await fetch(url, { cache: "no-store" });
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
      var res = await fetch(url, {
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

    // ---------------------------------------------------------------------
    // 필드 판별 규칙
    // ---------------------------------------------------------------------
    function isTextareaColumn(name) {
      return /비고|메모|설명|내용/.test(name);
    }
    function isDateColumn(name) {
      return /일자|날짜|시작일|종료일|체결일|계약일/.test(name);
    }
    function isNumberColumn(name) {
      return /용량|금액|단가|비율|수량|면적|사용량|발전량|REC|MW|MWh|kW|kWh|개월|년수/.test(name);
    }
    function isRequiredColumn(tableName, name) {
      var schema = SCHEMA_BY_KEY[tableName];
      return !!(schema && (name === schema.pk || (schema.fk && Object.prototype.hasOwnProperty.call(schema.fk, name))));
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

    function validateFieldLive(tableName, columnName, wrap, input) {
      var schema = SCHEMA_BY_KEY[tableName];
      var required = isRequiredColumn(tableName, columnName);
      var value = (input.value || "").trim();

      if (required && !value) {
        markInvalid(wrap, input, columnName + "는 필수입니다.");
        return false;
      }

      // 새 데이터 입력 중인데 PK가 이미 존재하는 값이면 - 막지는 않되(저장하면
      // 그 데이터를 덮어써 수정하는 것과 같으므로) 알려줍니다.
      if (schema && columnName === schema.pk && value && !formState.loadedPk) {
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
        if (!validateFieldLive(tableName, columnName, wrap, input)) ok = false;
      });
      return ok;
    }

    function createField(tableName, columnName) {
      var wrap = el("div", { class: "ppaf-row" });
      var labelClass = isRequiredColumn(tableName, columnName) ? "ppaf-label ppaf-required" : "ppaf-label";
      var label = el("label", { class: labelClass }, [columnName]);
      var fieldName = "fld_" + columnName;
      var input;

      var lookupTable = getLookupTable(tableName, columnName);
      if (lookupTable && (optionCache[lookupTable] || []).length > 0) {
        input = buildSelect(fieldName, optionCache[lookupTable] || []);
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

      var onLiveCheck = function () { validateFieldLive(tableName, columnName, wrap, input); };
      input.addEventListener("blur", onLiveCheck);
      input.addEventListener("change", onLiveCheck);
      input.addEventListener("input", function () {
        if (wrap.classList.contains("invalid")) onLiveCheck();
      });

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
        inp.value = record[key] || "";
      });
      fields.querySelectorAll(".ppaf-row").forEach(clearFieldNote);
    }

    function clearRecord() {
      fields.querySelectorAll("[data-name^='fld_']").forEach(function (inp) { inp.value = ""; });
      fields.querySelectorAll(".ppaf-row").forEach(clearFieldNote);
      formState.loadedPk = null;
      updateDeleteButtonState();
      var loadSel = document.getElementById("ppaf-load-select");
      if (loadSel) loadSel.value = "";
    }

    function updateDeleteButtonState() {
      deleteBtn.disabled = !formState.loadedPk;
      deleteBtn.title = formState.loadedPk ? "" : "먼저 '불러오기'로 기존 데이터를 선택하세요.";
    }

    function renderToolbar(tableName) {
      var toolbar = el("div", { class: "ppaf-toolbar" });
      var loadSelect = el("select", { class: "ppaf-input", id: "ppaf-load-select" });
      loadSelect.appendChild(el("option", { value: "" }, ["기존 데이터 선택"]));

      (optionCache[tableName] || []).forEach(function (opt) {
        loadSelect.appendChild(el("option", { value: opt.value }, [opt.label]));
      });

      var loadBtn = el("button", { class: "ppaf-btn", type: "button" }, ["불러오기"]);
      var newBtn = el("button", { class: "ppaf-btn", type: "button" }, ["새 입력"]);

      toolbar.appendChild(loadSelect);
      toolbar.appendChild(loadBtn);
      toolbar.appendChild(newBtn);
      fields.appendChild(toolbar);

      loadBtn.addEventListener("click", async function () {
        var pkValue = loadSelect.value || "";
        if (!pkValue) {
          showToast("불러올 데이터를 선택하세요.", "error");
          return;
        }
        withBusy(loadBtn, "불러오는 중...", async function () {
          var data = await apiGet(
            "/api/record?table=" + encodeURIComponent(tableName) + "&pk=" + encodeURIComponent(pkValue)
          );
          fillRecord(data.record || {});
          formState.loadedPk = pkValue;
          updateDeleteButtonState();
        }).catch(function (e) {
          showToast("불러오기 실패: " + (e.message || e), "error");
        });
      });

      newBtn.addEventListener("click", clearRecord);
    }

    function renderFields(tableName) {
      var schema = SCHEMA_BY_KEY[tableName];
      var columns = (schema && schema.columns) || [];
      columns.forEach(function (col) {
        fields.appendChild(createField(tableName, col));
      });
    }

    async function renderMode(forceReload) {
      var tableName = currentTable();
      fields.innerHTML = "";
      formState.loadedPk = null;
      help.textContent = "저장하면 지금 열려 있는(또는 열려 있지 않으면 새로 열리는) 엑셀 파일에 바로 반영됩니다.";
      await prepareOptionsForTable(tableName, !!forceReload);
      renderToolbar(tableName);
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
        await renderMode(true);
      } catch (e) {
        showToast("화면 로딩 실패: " + (e.message || e), "error");
      }
    }

    function closeModal() {
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
        location.reload();
      } catch (e) {
        showToast("새로고침 실패: " + (e.message || e), "error");
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
          return apiPost("/api/save", { table: tableName, record: record });
        });
        showToast(data.message || "저장 완료", "success");

        delete optionCache[tableName];
        await renderMode(true);
        fillRecord(record);
        formState.loadedPk = record[pkName] || null;
        updateDeleteButtonState();

        var loadSel = document.getElementById("ppaf-load-select");
        if (loadSel) loadSel.value = record[pkName] || "";

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
          return apiPost("/api/delete", { table: tableName, pk: pkValue });
        });
        showToast(data.message || "삭제 완료", "success");
        closeDeleteConfirm();
        clearRecord();
        delete optionCache[tableName];
        await renderMode(true);
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
        ".ppaf-head{padding:18px 20px;border-bottom:1px solid var(--ppaf-line);display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;background:var(--ppaf-bg);border-radius:18px 18px 0 0}" +
        ".ppaf-title{font-size:16px;font-weight:800}" +
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
        ".ppaf-toolbar{display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:end;margin:8px 0 16px;grid-column:1 / -1}" +
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
        "@media(max-width:640px){.ppaf-modal{right:12px;bottom:72px;width:calc(100vw - 24px)}.ppaf-fields{grid-template-columns:1fr}.ppaf-toolbar{grid-template-columns:1fr}.ppaf-open{right:16px;bottom:16px;padding:12px 16px;font-size:13px}}"
    });

    var openBtn = el("button", { class: "ppaf-open", type: "button" }, ["✎ 간편 입력/저장"]);
    var backdrop = el("div", { class: "ppaf-backdrop" });
    var modal = el("div", { class: "ppaf-modal" });

    var closeBtn = el("button", { class: "ppaf-btn", type: "button" }, ["닫기"]);
    var clearBtn = el("button", { class: "ppaf-btn", type: "button" }, ["초기화"]);
    var rebuildBtn = el("button", { class: "ppaf-btn", type: "button" }, ["대시보드 새로고침"]);
    var deleteBtn = el("button", { class: "ppaf-btn danger", type: "button", disabled: "disabled" }, ["삭제"]);
    var saveBtn = el("button", { class: "ppaf-btn primary", type: "button" }, ["엑셀에 저장"]);

    var modeSel = el("select", { class: "ppaf-input" });

    var help = el("div", { class: "ppaf-help" });
    var bodyWrap = el("div", { class: "ppaf-body" });
    var fields = el("div", { class: "ppaf-fields" });

    modal.appendChild(el("div", { class: "ppaf-head" }, [el("div", { class: "ppaf-title" }, ["간편 입력/저장 (엑셀에 바로 반영)"]), closeBtn]));
    bodyWrap.appendChild(modeSel);
    bodyWrap.appendChild(help);
    bodyWrap.appendChild(fields);
    modal.appendChild(bodyWrap);
    modal.appendChild(el("div", { class: "ppaf-foot" }, [clearBtn, deleteBtn, rebuildBtn, saveBtn]));

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
    clearBtn.addEventListener("click", clearRecord);
    rebuildBtn.addEventListener("click", function () { rebuildDashboard().catch(console.error); });
    saveBtn.addEventListener("click", function () { saveCurrentTable().catch(console.error); });
    deleteBtn.addEventListener("click", function () { openDeleteConfirm().catch(console.error); });
    confirmCancelBtn.addEventListener("click", closeDeleteConfirm);
    confirmBackdrop.addEventListener("click", closeDeleteConfirm);
    confirmDeleteBtn.addEventListener("click", function () { performDelete().catch(console.error); });
    modeSel.addEventListener("change", function () {
      renderMode(true).catch(function (e) { showToast("목록 로딩 실패: " + (e.message || e), "error"); });
    });

    document.addEventListener("keydown", function (e) {
      if (e.key !== "Escape") return;
      if (confirmBox.classList.contains("show")) { closeDeleteConfirm(); return; }
      if (modal.classList.contains("show")) { closeModal(); }
    });

    console.log("[ppa] dashboard_form.js loaded");
  } catch (e) {
    console.error("[ppa] dashboard_form.js error:", e);
  }
})();
