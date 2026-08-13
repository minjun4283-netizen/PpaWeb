(() => {
  try {
    if (window.__PPA_FORM_LOADED__) return;
    window.__PPA_FORM_LOADED__ = true;

    function onReady(fn) {
      if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", fn, { once: true });
      } else {
        fn();
      }
    }

    function el(tag, attrs = {}, children = []) {
      const node = document.createElement(tag);
      Object.entries(attrs).forEach(([k, v]) => {
        if (v === undefined || v === null) return;
        if (k === "class") node.className = v;
        else if (k === "html") node.innerHTML = v;
        else node.setAttribute(k, v);
      });
      children.forEach(ch => {
        if (typeof ch === "string") node.appendChild(document.createTextNode(ch));
        else if (ch) node.appendChild(ch);
      });
      return node;
    }

    function getTable(tableKey) {
      if (!window.DATA || !Array.isArray(window.DATA.tables)) return null;
      return window.DATA.tables.find(t => t.key === tableKey) || null;
    }

    function isHelperColumn(name) {
      if (!name) return true;
      const bad = new Set([
        "PK중복", "PK공란", "조합중복", "열1"
      ]);
      if (bad.has(name)) return true;
      if (name.endsWith("참조")) return true;
      if (name.endsWith("공란")) return true;
      if (name.endsWith("중복")) return true;
      return false;
    }

    function getColumnNames(tableKey) {
      const table = getTable(tableKey);
      if (!table) return [];

      const cols = [];

      if (Array.isArray(table.columns) && table.columns.length) {
        table.columns.forEach(c => {
          if (typeof c === "string") cols.push(c);
          else if (c && typeof c === "object") cols.push(c.name || c.key || c.label || "");
        });
      }

      if ((!cols.length) && Array.isArray(table.rows)) {
        table.rows.forEach(r => {
          const cells = (r && r.cells) || {};
          Object.keys(cells).forEach(k => cols.push(k));
        });
      }

      const seen = new Set();
      return cols.filter(c => {
        if (!c || seen.has(c) || isHelperColumn(c)) return false;
        seen.add(c);
        return true;
      });
    }

    function getRowsFromData(tableKey) {
      const table = getTable(tableKey);
      return table && Array.isArray(table.rows) ? table.rows : [];
    }

    function makeOptions(rows, idKey, nameKey) {
      return rows
        .map(r => r.cells || {})
        .filter(c => c[idKey])
        .map(c => ({
          value: c[idKey],
          label: `${c[idKey]} | ${c[nameKey] || ""}`
        }));
    }

    async function apiPost(url, payload) {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      const data = await res.json();
      if (!res.ok || !data.ok) throw new Error(data.error || "저장 실패");
      return data;
    }

    onReady(() => {
      if (!document.body) return;

      const MODES = {
        plantPurchase: {
          label: "발전소 + 구매계약",
          help: "T_발전소, T_구매계약의 컬럼 전체 입력",
          sections: [
            { title: "발전소 정보", tableKey: "T_발전소", prefix: "plant_" },
            { title: "구매계약 정보", tableKey: "T_구매계약", prefix: "purchase_" }
          ]
        },
        demandBundle: {
          label: "수요기업 + 판매계약 + 전기사용지",
          help: "T_수요기업, T_판매계약, T_전기사용지의 컬럼 전체 입력",
          sections: [
            { title: "수요기업 정보", tableKey: "T_수요기업", prefix: "customer_" },
            { title: "판매계약 정보", tableKey: "T_판매계약", prefix: "sales_" },
            { title: "전기사용지 정보", tableKey: "T_전기사용지", prefix: "site_" }
          ]
        },
        matching: {
          label: "수급매칭",
          help: "T_수급매칭의 컬럼 전체 입력",
          sections: [
            { title: "수급매칭 정보", tableKey: "T_수급매칭", prefix: "matching_" }
          ]
        }
      };

      const AUTO_LINKS = {
        "purchase_발전소ID": "plant_발전소ID",
        "sales_수요기업ID": "customer_수요기업ID",
        "site_판매계약ID": "sales_판매계약ID"
      };

      const REF_SELECTS = {
        "matching_구매계약ID": () => makeOptions(getRowsFromData("T_구매계약"), "구매계약ID", "계약명"),
        "matching_전기사용지ID": () => makeOptions(getRowsFromData("T_전기사용지"), "전기사용지ID", "사용지명")
      };

      const style = el("style", {
        html: `
          .ppaf-open{position:fixed;right:24px;bottom:24px;z-index:2147483647;border:none;border-radius:999px;padding:12px 16px;background:#0B8577;color:#fff;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(0,0,0,.18)}
          .ppaf-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2147483645;display:none}
          .ppaf-backdrop.show{display:block}
          .ppaf-modal{position:fixed;right:24px;bottom:80px;width:min(1200px,calc(100vw - 32px));max-height:84vh;overflow:auto;background:#fff;border:1px solid #d9e5e1;border-radius:16px;box-shadow:0 16px 40px rgba(0,0,0,.22);z-index:2147483646;display:none}
          .ppaf-modal.show{display:block}
          .ppaf-head{padding:16px 18px;border-bottom:1px solid #e6efec;display:flex;justify-content:space-between;align-items:center}
          .ppaf-title{font-size:16px;font-weight:800}
          .ppaf-body{padding:16px 18px}
          .ppaf-help{font-size:12px;color:#5b6b65;margin:8px 0 14px}
          .ppaf-fields{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}
          .ppaf-row{display:flex;flex-direction:column;gap:4px}
          .ppaf-label{font-size:12px;font-weight:700;color:#33413d}
          .ppaf-input,.ppaf-textarea{font-size:13px;padding:9px 10px;border:1px solid #d9e5e1;border-radius:8px}
          .ppaf-textarea{min-height:88px;resize:vertical}
          .ppaf-foot{padding:14px 18px;border-top:1px solid #e6efec;display:flex;justify-content:flex-end;gap:8px}
          .ppaf-btn{border:1px solid #d9e5e1;background:#fff;border-radius:8px;padding:9px 12px;cursor:pointer;font-weight:700}
          .ppaf-btn.primary{background:#0B8577;border-color:#0B8577;color:#fff}
          .ppaf-sec{grid-column:1 / -1;margin:14px 0 4px;font-weight:800;color:#0B8577;border-top:1px solid #e6efec;padding-top:10px}
          .ppaf-readonly{background:#f5f7f7}
        `
      });

      const openBtn = el("button", { class: "ppaf-open", type: "button" }, ["간편 입력/저장"]);
      const backdrop = el("div", { class: "ppaf-backdrop" });
      const modal = el("div", { class: "ppaf-modal" });

      const closeBtn = el("button", { class: "ppaf-btn", type: "button" }, ["닫기"]);
      const clearBtn = el("button", { class: "ppaf-btn", type: "button" }, ["초기화"]);
      const saveBtn = el("button", { class: "ppaf-btn primary", type: "button" }, ["저장"]);

      const modeSel = el("select", { class: "ppaf-input" });
      Object.entries(MODES).forEach(([k, v]) => {
        modeSel.appendChild(el("option", { value: k }, [v.label]));
      });

      const help = el("div", { class: "ppaf-help" });
      const fields = el("div", { class: "ppaf-fields" });

      modal.appendChild(
        el("div", { class: "ppaf-head" }, [
          el("div", { class: "ppaf-title" }, ["간편 입력/저장"]),
          closeBtn
        ])
      );
      modal.appendChild(el("div", { class: "ppaf-body" }, [modeSel, help, fields]));
      modal.appendChild(el("div", { class: "ppaf-foot" }, [clearBtn, saveBtn]));

      function createField(fullName, columnName) {
        const wrap = el("div", { class: "ppaf-row" });
        const label = el("label", { class: "ppaf-label" }, [columnName]);
        const refOptionsFactory = REF_SELECTS[fullName];
        const linkedFrom = AUTO_LINKS[fullName];

        let input;

        if (refOptionsFactory) {
          input = el("select", { class: "ppaf-input", "data-name": fullName });
          input.appendChild(el("option", { value: "" }, ["선택"]));
          refOptionsFactory().forEach(o => {
            input.appendChild(el("option", { value: o.value }, [o.label]));
          });
        } else if (/비고|메모|설명|내용/.test(columnName)) {
          input = el("textarea", { class: "ppaf-textarea", "data-name": fullName });
        } else {
          input = el("input", { class: "ppaf-input", type: "text", "data-name": fullName });
        }

        if (linkedFrom) {
          input.readOnly = true;
          input.classList.add("ppaf-readonly");
        }

        wrap.appendChild(label);
        wrap.appendChild(input);
        return wrap;
      }

      function renderSection(section) {
        fields.appendChild(el("div", { class: "ppaf-sec" }, [section.title]));
        const columns = getColumnNames(section.tableKey);

        columns.forEach(col => {
          fields.appendChild(createField(section.prefix + col, col));
        });
      }

      function syncAutoLinks() {
        Object.entries(AUTO_LINKS).forEach(([targetName, sourceName]) => {
          const source = fields.querySelector(`[data-name="${sourceName}"]`);
          const target = fields.querySelector(`[data-name="${targetName}"]`);
          if (!source || !target) return;

          const apply = () => {
            target.value = source.value || "";
          };

          source.addEventListener("input", apply);
          source.addEventListener("change", apply);
          apply();
        });
      }

      function collectByPrefix(prefix) {
        const obj = {};
        fields.querySelectorAll(`[data-name^="${prefix}"]`).forEach(inp => {
          const key = inp.getAttribute("data-name").replace(prefix, "");
          obj[key] = inp.value || "";
        });
        return obj;
      }

      function renderMode() {
        const mode = modeSel.value;
        const conf = MODES[mode];
        fields.innerHTML = "";
        help.textContent = conf.help;

        conf.sections.forEach(renderSection);
        syncAutoLinks();
      }

      function clearFields() {
        fields.querySelectorAll("[data-name]").forEach(inp => {
          inp.value = "";
        });
        syncAutoLinks();
      }

      function openModal() {
        backdrop.classList.add("show");
        modal.classList.add("show");
      }

      function closeModal() {
        backdrop.classList.remove("show");
        modal.classList.remove("show");
      }

      async function saveByMode() {
        saveBtn.disabled = true;
        saveBtn.textContent = "저장 중...";

        try {
          const mode = modeSel.value;
          let data;

          if (mode === "plantPurchase") {
            data = await apiPost("/api/save-plant-purchase", {
              plant: collectByPrefix("plant_"),
              purchase: collectByPrefix("purchase_")
            });
          } else if (mode === "demandBundle") {
            data = await apiPost("/api/save-demand-sales-site", {
              customer: collectByPrefix("customer_"),
              sales: collectByPrefix("sales_"),
              site: collectByPrefix("site_")
            });
          } else {
            data = await apiPost("/api/save-matching", {
              matching: collectByPrefix("matching_")
            });
          }

          alert((data.message || "저장 완료") + "\n\n약 8초 후 새로고침합니다.");
          closeModal();
          setTimeout(() => location.reload(), 8000);
        } catch (e) {
          alert("저장 실패\n" + (e.message || e));
        } finally {
          saveBtn.disabled = false;
          saveBtn.textContent = "저장";
        }
      }

      document.head.appendChild(style);
      document.body.appendChild(backdrop);
      document.body.appendChild(modal);
      document.body.appendChild(openBtn);

      openBtn.addEventListener("click", openModal);
      closeBtn.addEventListener("click", closeModal);
      backdrop.addEventListener("click", closeModal);
      clearBtn.addEventListener("click", clearFields);
      saveBtn.addEventListener("click", saveByMode);
      modeSel.addEventListener("change", renderMode);

      renderMode();
      console.log("[ppa] full-column dashboard_form.js loaded");
    });
  } catch (e) {
    console.error("[ppa] dashboard_form.js error:", e);
  }
})();