(function () {
  try {
    if (window.__PPA_FORM_LOADED__) return;
    window.__PPA_FORM_LOADED__ = true;

    var SCHEMA = {
      "T_발전소": ["발전소ID", "발전소명", "발전법인명", "설비용량(MW)", "발전원", "Readiness", "MGA_Supply"],
      "T_구매계약": ["구매계약ID", "발전소ID", "구매계약용량(MW)", "구매단가(원/kWh)", "공급기한_구매", "계약기간(년)", "수요기업 미확보", "구매 담당자"],
      "T_수요기업": ["수요기업ID", "기업명"],
      "T_판매계약": ["판매계약ID", "수요기업ID", "판매계약용량(MW)", "계약일", "공급기한_판매", "계약유형", "판매단가(원/kWh)", "공급자원 미확보", "판매 담당자", "계약기간(년)", "Requirement", "MGA_Demand"],
      "T_전기사용지": ["전기사용지ID", "판매계약ID", "전기사용지명", "전기사용지계약용량(MW)"],
      "T_수급매칭": ["수급매칭ID", "전기사용지ID", "구매계약ID", "현황"]
    };

    var PK_BY_TABLE = {
      "T_발전소": "발전소ID",
      "T_구매계약": "구매계약ID",
      "T_수요기업": "수요기업ID",
      "T_판매계약": "판매계약ID",
      "T_전기사용지": "전기사용지ID",
      "T_수급매칭": "수급매칭ID"
    };

    var TABLE_META = {
      "T_발전소": { label: "발전소", help: "발전소 데이터를 신규 등록하거나 기존 데이터 수정 후 저장합니다." },
      "T_구매계약": { label: "구매계약", help: "구매계약 데이터를 신규 등록하거나 기존 데이터 수정 후 저장합니다." },
      "T_수요기업": { label: "수요기업", help: "수요기업 데이터를 신규 등록하거나 기존 데이터 수정 후 저장합니다." },
      "T_판매계약": { label: "판매계약", help: "판매계약 데이터를 신규 등록하거나 기존 데이터 수정 후 저장합니다." },
      "T_전기사용지": { label: "전기사용지", help: "전기사용지 데이터를 신규 등록하거나 기존 데이터 수정 후 저장합니다." },
      "T_수급매칭": { label: "수급매칭", help: "수급매칭 데이터를 신규 등록하거나 기존 데이터 수정 후 저장합니다." }
    };

    var FIELD_LOOKUP = {
      "T_구매계약": { "발전소ID": "T_발전소" },
      "T_판매계약": { "수요기업ID": "T_수요기업" },
      "T_전기사용지": { "판매계약ID": "T_판매계약" },
      "T_수급매칭": { "전기사용지ID": "T_전기사용지", "구매계약ID": "T_구매계약" }
    };

    var optionCache = {};

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

      if (!data.ok) {
        throw new Error(data.error || "조회 실패");
      }

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

      if (!data.ok) {
        throw new Error(data.error || "요청 실패");
      }

      return data;
    }

    function isTextareaColumn(name) {
      return /비고|메모|설명|내용/.test(name);
    }

    function isDateColumn(name) {
      return /일자|날짜|시작일|종료일|체결일|계약일/.test(name);
    }

    function isNumberColumn(name) {
      return /용량|금액|단가|비율|수량|면적|사용량|발전량|REC|MW|MWh|kW|kWh|개월|년수/.test(name);
    }

    function isRequiredColumn(name) {
      return /ID$/.test(name);
    }

    function currentTable() {
      return modeSel.value;
    }

    async function getTableOptions(tableName, forceReload) {
      if (!forceReload && optionCache[tableName]) {
        return optionCache[tableName];
      }
      var data = await apiGet("/api/list-records?table=" + encodeURIComponent(tableName));
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

    function createField(tableName, columnName) {
      var wrap = el("div", { class: "ppaf-row" });
      var labelClass = isRequiredColumn(columnName) ? "ppaf-label ppaf-required" : "ppaf-label";
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
    }

    function clearRecord() {
      fields.querySelectorAll("[data-name^='fld_']").forEach(function (inp) {
        inp.value = "";
      });

      var loadSel = document.getElementById("ppaf-load-select");
      if (loadSel) loadSel.value = "";
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
          alert("불러올 데이터를 선택하세요.");
          return;
        }

        try {
          var data = await apiGet(
            "/api/get-record?table=" +
            encodeURIComponent(tableName) +
            "&pk_value=" +
            encodeURIComponent(pkValue)
          );
          fillRecord(data.record || {});
        } catch (e) {
          alert("불러오기 실패\n" + (e.message || e));
        }
      });

      newBtn.addEventListener("click", function () {
        clearRecord();
      });
    }

    function renderFields(tableName) {
      var columns = SCHEMA[tableName] || [];
      columns.forEach(function (col) {
        fields.appendChild(createField(tableName, col));
      });
    }

    async function renderMode(forceReload) {
      var tableName = currentTable();
      fields.innerHTML = "";
      help.textContent = TABLE_META[tableName].help || "";

      await prepareOptionsForTable(tableName, !!forceReload);
      renderToolbar(tableName);
      renderFields(tableName);
    }

    async function openModal() {
      backdrop.classList.add("show");
      modal.classList.add("show");
      try {
        await renderMode(true);
      } catch (e) {
        alert("화면 로딩 실패\n" + (e.message || e));
      }
    }

    function closeModal() {
      backdrop.classList.remove("show");
      modal.classList.remove("show");
    }

    async function rebuildDashboard() {
      rebuildBtn.disabled = true;
      rebuildBtn.textContent = "재생성 중...";
      try {
        var data = await apiPost("/api/rebuild", {});
        alert(data.message || "재생성을 시작했습니다.");
      } catch (e) {
        alert("재생성 실패\n" + (e.message || e));
      } finally {
        rebuildBtn.disabled = false;
        rebuildBtn.textContent = "대시보드 재생성";
      }
    }

    async function saveCurrentTable() {
      var tableName = currentTable();
      var record = collectRecord();
      var pkName = PK_BY_TABLE[tableName];

      if (!record[pkName]) {
        alert(pkName + "는 필수입니다.");
        return;
      }

      saveBtn.disabled = true;
      saveBtn.textContent = "저장 중...";

      try {
        var data = await apiPost("/api/save-table", {
          table: tableName,
          record: record
        });

        alert(data.message || "저장 완료");

        delete optionCache[tableName];
        await renderMode(true);
        fillRecord(record);

        var loadSel = document.getElementById("ppaf-load-select");
        if (loadSel) loadSel.value = record[pkName] || "";
      } catch (e) {
        alert("저장 실패\n" + (e.message || e));
      } finally {
        saveBtn.disabled = false;
        saveBtn.textContent = "저장";
      }
    }

    var style = el("style", {
      html:
        ".ppaf-open{position:fixed;right:24px;bottom:24px;z-index:2147483647;border:none;border-radius:999px;padding:12px 16px;background:#0B8577;color:#fff;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(0,0,0,.18)}" +
        ".ppaf-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2147483645;display:none}" +
        ".ppaf-backdrop.show{display:block}" +
        ".ppaf-modal{position:fixed;right:24px;bottom:80px;width:min(1100px,calc(100vw - 32px));max-height:84vh;overflow:auto;background:#fff;border:1px solid #d9e5e1;border-radius:16px;box-shadow:0 16px 40px rgba(0,0,0,.22);z-index:2147483646;display:none}" +
        ".ppaf-modal.show{display:block}" +
        ".ppaf-head{padding:16px 18px;border-bottom:1px solid #e6efec;display:flex;justify-content:space-between;align-items:center}" +
        ".ppaf-title{font-size:16px;font-weight:800}" +
        ".ppaf-body{padding:16px 18px}" +
        ".ppaf-help{font-size:12px;color:#5b6b65;margin:8px 0 14px}" +
        ".ppaf-fields{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}" +
        ".ppaf-row{display:flex;flex-direction:column;gap:4px}" +
        ".ppaf-label{font-size:12px;font-weight:700;color:#33413d}" +
        ".ppaf-input,.ppaf-textarea{font-size:13px;padding:9px 10px;border:1px solid #d9e5e1;border-radius:8px}" +
        ".ppaf-textarea{min-height:88px;resize:vertical}" +
        ".ppaf-foot{padding:14px 18px;border-top:1px solid #e6efec;display:flex;justify-content:flex-end;gap:8px}" +
        ".ppaf-btn{border:1px solid #d9e5e1;background:#fff;border-radius:8px;padding:9px 12px;cursor:pointer;font-weight:700}" +
        ".ppaf-btn.primary{background:#0B8577;border-color:#0B8577;color:#fff}" +
        ".ppaf-required::after{content:' *';color:#d92d20;font-weight:800}" +
        ".ppaf-toolbar{display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:end;margin:8px 0 14px;grid-column:1 / -1}" +
        "@media(max-width:640px){.ppaf-modal{right:12px;bottom:72px;width:calc(100vw - 24px)}.ppaf-fields{grid-template-columns:1fr}.ppaf-toolbar{grid-template-columns:1fr}}"
    });

    var openBtn = el("button", { class: "ppaf-open", type: "button" }, ["간편 입력/저장"]);
    var backdrop = el("div", { class: "ppaf-backdrop" });
    var modal = el("div", { class: "ppaf-modal" });

    var closeBtn = el("button", { class: "ppaf-btn", type: "button" }, ["닫기"]);
    var clearBtn = el("button", { class: "ppaf-btn", type: "button" }, ["초기화"]);
    var rebuildBtn = el("button", { class: "ppaf-btn", type: "button" }, ["대시보드 재생성"]);
    var saveBtn = el("button", { class: "ppaf-btn primary", type: "button" }, ["저장"]);

    var modeSel = el("select", { class: "ppaf-input" });
    Object.keys(TABLE_META).forEach(function (tableName) {
      modeSel.appendChild(el("option", { value: tableName }, [TABLE_META[tableName].label]));
    });

    var help = el("div", { class: "ppaf-help" });
    var bodyWrap = el("div", { class: "ppaf-body" });
    var fields = el("div", { class: "ppaf-fields" });

    modal.appendChild(
      el("div", { class: "ppaf-head" }, [
        el("div", { class: "ppaf-title" }, ["간편 입력/저장"]),
        closeBtn
      ])
    );

    bodyWrap.appendChild(modeSel);
    bodyWrap.appendChild(help);
    bodyWrap.appendChild(fields);

    modal.appendChild(bodyWrap);
    modal.appendChild(el("div", { class: "ppaf-foot" }, [clearBtn, rebuildBtn, saveBtn]));

    document.head.appendChild(style);
    document.body.appendChild(backdrop);
    document.body.appendChild(modal);
    document.body.appendChild(openBtn);

    openBtn.addEventListener("click", function () {
      openModal().catch(console.error);
    });
    closeBtn.addEventListener("click", closeModal);
    backdrop.addEventListener("click", closeModal);
    clearBtn.addEventListener("click", clearRecord);
    rebuildBtn.addEventListener("click", function () {
      rebuildDashboard().catch(console.error);
    });
    saveBtn.addEventListener("click", function () {
      saveCurrentTable().catch(console.error);
    });
    modeSel.addEventListener("change", function () {
      renderMode(true).catch(function (e) {
        alert("목록 로딩 실패\n" + (e.message || e));
      });
    });

    console.log("[ppa] dashboard_form.js loaded");
  } catch (e) {
    console.error("[ppa] dashboard_form.js error:", e);
  }
})();