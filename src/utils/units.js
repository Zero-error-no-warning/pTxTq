export const EMU_PER_INCH = 914400;
export const PT_PER_INCH = 72;
export const PX_PER_INCH = 96;

export function emuToPt(emu) {
  return (Number(emu) || 0) * (PT_PER_INCH / EMU_PER_INCH);
}

export function ptToEmu(pt) {
  return Math.round((Number(pt) || 0) * (EMU_PER_INCH / PT_PER_INCH));
}

export function emuToPx(emu) {
  return (Number(emu) || 0) * (PX_PER_INCH / EMU_PER_INCH);
}

export function pxToEmu(px) {
  return Math.round((Number(px) || 0) * (EMU_PER_INCH / PX_PER_INCH));
}

export function centipointsToPt(value) {
  return (Number(value) || 0) / 100;
}

export function ptToCentipoints(value) {
  return Math.round((Number(value) || 0) * 100);
}

export function deg60000ToDeg(value) {
  return (Number(value) || 0) / 60000;
}

export function degToDeg60000(value) {
  return Math.round((Number(value) || 0) * 60000);
}

export function emuToCssPxString(emu) {
  return `${emuToPx(emu)}px`;
}
