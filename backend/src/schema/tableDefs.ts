// Seed schema for the 6 PPA tables the original VBA macro managed
// (T_발전소, T_구매계약, T_수급매칭, T_전기사용지, T_판매계약, T_수요기업).
//
// T_수요기업, T_판매계약, T_전기사용지, T_구매계약, T_수급매칭 below are the
// REAL columns confirmed from screenshots of the actual xlsm (validation
// flag columns like PK중복/PK공란/*참조 are excluded here since this system
// computes those on demand instead of storing them). T_발전소 is still a
// placeholder guess pending its own screenshot — fix it the same way once
// that's available.

export type ColumnType = "text" | "number" | "date" | "boolean";

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
      { key: "발전법인명", label: "발전법인명", type: "text" },
      { key: "설비용량_MW", label: "설비용량(MW)", type: "number" },
      { key: "발전원", label: "발전원", type: "text" },
      { key: "Readiness", label: "Readiness", type: "text" },
      { key: "MGA_Supply", label: "MGA_Supply", type: "number" },
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
      { key: "구매계약용량_MW", label: "구매계약용량(MW)", type: "number" },
      { key: "구매단가", label: "구매단가(원/kWh)", type: "number" },
      { key: "공급기한_구매", label: "공급기한_구매", type: "date" },
      { key: "계약기간_년", label: "계약기간(년)", type: "number" },
      { key: "수요기업_미확보", label: "수요기업 미확보", type: "boolean" },
      { key: "구매_담당자", label: "구매 담당자", type: "text" },
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
      { key: "판매계약용량_MW", label: "판매계약용량(MW)", type: "number" },
      { key: "계약일", label: "계약일", type: "date" },
      { key: "공급기한_판매", label: "공급기한_판매", type: "date" },
      { key: "계약유형", label: "계약유형", type: "text" },
      { key: "판매단가", label: "판매단가(원/kWh)", type: "number" },
      { key: "공급자원_미확보", label: "공급자원 미확보", type: "boolean" },
      { key: "판매_담당자", label: "판매 담당자", type: "text" },
      { key: "계약기간_년", label: "계약기간(년)", type: "number" },
      { key: "Requirement", label: "Requirement", type: "text" },
      { key: "MGA_Demand", label: "MGA_Demand", type: "number" },
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
      { key: "전기사용지명", label: "전기사용지명", type: "text" },
      { key: "전기사용지계약용량_MW", label: "전기사용지계약용량(MW)", type: "number" },
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
      { key: "현황", label: "현황", type: "text" },
    ],
    foreignKeys: [
      { column: "전기사용지ID", refTable: "T_전기사용지", refColumn: "전기사용지ID" },
      { column: "구매계약ID", refTable: "T_구매계약", refColumn: "구매계약ID" },
    ],
    uniqueGroups: [["전기사용지ID", "구매계약ID"]],
  },
];
