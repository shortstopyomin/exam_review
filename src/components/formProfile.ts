"use client";

import { useMemo, useSyncExternalStore } from "react";

/*
 * 請假單、外出單共用：姓名、單位、職務、聯繫方式記在這台瀏覽器，下次不用重打。
 * 只存在 localStorage（不會上傳）；無痕模式或被封鎖時就不記。
 */

const PROFILE_KEY = "leave-form-profile";

export interface Profile {
  name?: string;
  dept?: string;
  position?: string;
  contact?: string;
}

const noSubscribe = () => () => {};

function readRaw(): string | null {
  try {
    return localStorage.getItem(PROFILE_KEY);
  } catch {
    return null;
  }
}

function parse(raw: string | null): Profile {
  try {
    const p = JSON.parse(raw ?? "{}");
    const str = (v: unknown) => (typeof v === "string" ? v : undefined);
    return { name: str(p.name), dept: str(p.dept), position: str(p.position), contact: str(p.contact) };
  } catch {
    return {};
  }
}

/** 上次記住的資料（伺服器端產生頁面時是空的） */
export function useStoredProfile(): Profile {
  const raw = useSyncExternalStore(noSubscribe, readRaw, () => null);
  return useMemo(() => parse(raw), [raw]);
}

/** 記住這次填的資料；只更新有給的欄位，請假單不會清掉外出單記的聯繫方式 */
export function saveProfile(update: Profile) {
  try {
    localStorage.setItem(PROFILE_KEY, JSON.stringify({ ...parse(readRaw()), ...update }));
  } catch {
    /* 記不住也沒關係 */
  }
}

/** 今天的日期 yyyy-mm-dd（本地時間） */
export function todayIso() {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
