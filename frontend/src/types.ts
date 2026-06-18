export type SlideStatus = "changed" | "unchanged" | "added" | "removed";

export interface ChangeSpan {
  type: "changed" | "added" | "removed";
  old: string;
  new: string;
}

export interface ChangeTag {
  type: string;
  detail: string;
}

export interface SlideResult {
  old_slide: number | null;
  new_slide: number | null;
  status: SlideStatus;
  possibly_reordered: boolean;
  has_changes: boolean;
  has_notes_changes: boolean;
  text_similarity: number;
  changes: ChangeSpan[];
  old_text: string;
  new_text: string;
  old_notes: string;
  new_notes: string;
  summary: string;
  change_tags: ChangeTag[];
  ai_summary: boolean;
  old_image: string | null;
  new_image: string | null;
}

export interface CompareSummary {
  total_old: number;
  total_new: number;
  changed: number;
  unchanged: number;
  added: number;
  removed: number;
}

export interface CompareResponse {
  summary: CompareSummary;
  slides: SlideResult[];
  warnings: string[];
}
