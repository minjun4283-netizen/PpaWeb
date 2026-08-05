// Seed schema for the 6 PPA tables the original VBA macro managed
// (T_발전소, T_구매계약, T_수급매칭, T_전기사용지, T_판매계약, T_수요기업).
//
// IMPORTANT: the VBA source only reveals the PK columns and a handful of
// columns touched by validation/tooltip logic. The remaining business
// columns below (계약기간, 단가, 용량 등) are a reasonable domain guess for a
// renewable-energy PPA contract system, NOT the real spreadsheet headers.
// Adjust them from the 컬럼관리 admin screen (or the column_defs table)
// once the real xlsm headers are available.

export type ColumnType = "text" | "number" | "date";

export interface ColumnSeed {
  key: string;
  label: string;
  type: ColumnType;
}

export interface ForeignKeySeed {
  column: string;
  refTable: string;
  refColumn: string;
}

export interface TableSeed {
  key: string;
  label: string;
  pk: string;
  columns: ColumnSeed[];
  foreignKeys: ForeignKeySeed[];
  // Groups of columns that together must be unique (mirrors ValidateCombinationDuplicate).
  uniqueGroups: string[][];
}

// Ordered as a natural data-entry flow: build up the supply chain
// (발전소→구매계약) and the demand chain (수요기업→판매계약→전기사용지)
// independently, then join them last with 수급매칭. This drives both the
// sidebar navigation order and the 검증 리포트's table grouping order.
export const TABLE_SEEDS: TableSeed[] = [
  {
    key: "T_발전소",
    label: "발전소",
    pk: "발전소ID",
    columns: [
      { key: "발전소ID", label: "발전소ID", type: "text" },
      { key: "발전소명", label: "발전소명", type: "text" },
      { key: "발전소위치", label: "발전소위치", type: "text" },
      { key: "발전원유형", label: "발전원유형", type: "text" },
      { key: "설비용량_kW", label: "설비용량(kW)", type: "number" },
      { key: "사업자명", label: "사업자명", type: "text" },
      { key: "비고", label: "비고", type: "text" },
    ],
    foreignKeys: [],
    uniqueGroups: [],
  },
  {
    key: "T_구매계약",
    label: "구매계약",
    pk: "구매계약ID",
    columns: [
      { key: "구매계약ID", label: "구매계약ID", type: "text" },
      { key: "발전소ID", label: "발전소ID", type: "text" },
      { key: "계약시작일", label: "계약시작일", type: "date" },
      { key: "계약종료일", label: "계약종료일", type: "date" },
      { key: "구매단가", label: "구매단가(원/kWh)", type: "number" },
      { key: "계약상태", label: "계약상태", type: "text" },
      { key: "비고", label: "비고", type: "text" },
    ],
    foreignKeys: [{ column: "발전소ID", refTable: "T_발전소", refColumn: "발전소ID" }],
    uniqueGroups: [],
  },
  {
    key: "T_수요기업",
    label: "수요기업",
    pk: "수요기업ID",
    columns: [
      { key: "수요기업ID", label: "수요기업ID", type: "text" },
      { key: "기업명", label: "기업명", type: "text" },
      { key: "사업자등록번호", label: "사업자등록번호", type: "text" },
      { key: "담당자", label: "담당자", type: "text" },
      { key: "연락처", label: "연락처", type: "text" },
      { key: "비고", label: "비고", type: "text" },
    ],
    foreignKeys: [],
    uniqueGroups: [],
  },
  {
    key: "T_판매계약",
    label: "판매계약",
    pk: "판매계약ID",
    columns: [
      { key: "판매계약ID", label: "판매계약ID", type: "text" },
      { key: "수요기업ID", label: "수요기업ID", type: "text" },
      { key: "계약시작일", label: "계약시작일", type: "date" },
      { key: "계약종료일", label: "계약종료일", type: "date" },
      { key: "판매단가", label: "판매단가(원/kWh)", type: "number" },
      { key: "계약상태", label: "계약상태", type: "text" },
      { key: "비고", label: "비고", type: "text" },
    ],
    foreignKeys: [{ column: "수요기업ID", refTable: "T_수요기업", refColumn: "수요기업ID" }],
    uniqueGroups: [],
  },
  {
    key: "T_전기사용지",
    label: "전기사용지",
    pk: "전기사용지ID",
    columns: [
      { key: "전기사용지ID", label: "전기사용지ID", type: "text" },
      { key: "판매계약ID", label: "판매계약ID", type: "text" },
      { key: "사용지명", label: "사용지명", type: "text" },
      { key: "주소", label: "주소", type: "text" },
      { key: "계약전력_kW", label: "계약전력(kW)", type: "number" },
      { key: "비고", label: "비고", type: "text" },
    ],
    foreignKeys: [{ column: "판매계약ID", refTable: "T_판매계약", refColumn: "판매계약ID" }],
    uniqueGroups: [],
  },
  {
    key: "T_수급매칭",
    label: "수급매칭",
    pk: "수급매칭ID",
    columns: [
      { key: "수급매칭ID", label: "수급매칭ID", type: "text" },
      { key: "전기사용지ID", label: "전기사용지ID", type: "text" },
      { key: "구매계약ID", label: "구매계약ID", type: "text" },
      { key: "매칭비율", label: "매칭비율(%)", type: "number" },
      { key: "적용시작일", label: "적용시작일", type: "date" },
      { key: "적용종료일", label: "적용종료일", type: "date" },
      { key: "비고", label: "비고", type: "text" },
    ],
    foreignKeys: [
      { column: "전기사용지ID", refTable: "T_전기사용지", refColumn: "전기사용지ID" },
      { column: "구매계약ID", refTable: "T_구매계약", refColumn: "구매계약ID" },
    ],
    uniqueGroups: [["전기사용지ID", "구매계약ID"]],
  },
];

// Tooltip field candidates, mirroring GetRelatedValueFromTableSafe's header fallback lists.
export const TOOLTIP_FIELD_CANDIDATES: Record<string, string[]> = {
  발전소명: ["발전소명", "발전소명칭", "시설명"],
  발전원유형: ["발전원유형", "발전원", "에너지원", "발전유형"],
  설비용량: ["설비용량_kW", "설비용량", "용량"],
  기업명: ["기업명", "수요기업명", "회사명"],
};
