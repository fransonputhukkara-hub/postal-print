export interface School {
  id: string;
  name: string;
  address: string;
  selected: boolean;
}

export interface SenderInfo {
  id: string;
  name: string;
  address: string;
  logoUrl?: string; // base64 encoded data URL for manual logo upload
}

export interface PrinterPreset {
  id: string;
  name: string;
  toBlockTopPadding: number;
  toBlockXShift: number;
  toBlockWidth: number;
  isBuiltIn?: boolean;
}

