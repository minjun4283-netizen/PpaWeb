// Ports the T_수급매칭 tooltip feature (Build전기사용지Tooltip_Expanded /
// Build구매계약Tooltip_Expanded) from a VBA UserForm popup into a plain JSON
// lookup the frontend renders as a hover/tap card.
import { dataTableName, db, quoteIdent } from "../db.js";

function getRelatedValue(tableKey: string, keyColumn: string, keyValue: string, returnColumn: string): string {
  if (!keyValue) return "";
  const physical = dataTableName(tableKey);
  const row = db
    .prepare(
      `SELECT ${quoteIdent(returnColumn)} as v FROM ${quoteIdent(physical)} WHERE TRIM(${quoteIdent(keyColumn)}) = ?`
    )
    .get(keyValue) as { v: unknown } | undefined;
  return row?.v != null ? String(row.v) : "";
}

export interface TooltipField {
  label: string;
  value: string;
}

export function buildElectricUseSiteTooltip(electricUseSiteId: string): TooltipField[] {
  const fields: TooltipField[] = [{ label: "전기사용지ID", value: electricUseSiteId }];

  const saleContractId = getRelatedValue("T_전기사용지", "전기사용지ID", electricUseSiteId, "판매계약ID");
  fields.push({ label: "판매계약ID", value: saleContractId || "(연결값 없음)" });
  if (!saleContractId) return fields;

  const demandCompanyId = getRelatedValue("T_판매계약", "판매계약ID", saleContractId, "수요기업ID");
  fields.push({ label: "수요기업ID", value: demandCompanyId || "(연결값 없음)" });

  const companyName = demandCompanyId
    ? getRelatedValue("T_수요기업", "수요기업ID", demandCompanyId, "기업명")
    : "";
  fields.push({ label: "기업명", value: companyName || "(연결값 없음)" });

  return fields;
}

export function buildBuyContractTooltip(buyContractId: string): TooltipField[] {
  const fields: TooltipField[] = [{ label: "구매계약ID", value: buyContractId }];

  const plantId = getRelatedValue("T_구매계약", "구매계약ID", buyContractId, "발전소ID");
  fields.push({ label: "발전소ID", value: plantId || "(연결값 없음)" });
  if (!plantId) return fields;

  const plantName = getRelatedValue("T_발전소", "발전소ID", plantId, "발전소명");
  fields.push({ label: "발전소명", value: plantName || "(연결값 없음)" });

  const plantType = getRelatedValue("T_발전소", "발전소ID", plantId, "발전원유형");
  fields.push({ label: "발전원 유형", value: plantType || "(값 없음)" });

  const plantCapacity = getRelatedValue("T_발전소", "발전소ID", plantId, "설비용량_kW");
  fields.push({ label: "설비용량(kW)", value: plantCapacity || "(값 없음)" });

  return fields;
}
