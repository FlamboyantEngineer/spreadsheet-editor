import type ExcelJS from 'exceljs';

// ECMA-376 indexed palette. Theme indices use light/dark order, unlike clrScheme XML.
const palette = `000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF
000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF
800000 008000 000080 808000 800080 008080 C0C0C0 808080
9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF
000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF
00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99
3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696
003366 339966 003300 333300 993300 993366 333399 333333`.split(/\s+/);
export function readThemeColors(xml: string): string[] {
  return [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ].map((name) => {
    const part = new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`).exec(xml)?.[1] ?? '';
    return /\b(?:lastClr|val)="([\da-f]{6})"/i.exec(part)?.[1] ?? '';
  });
}
type SpreadsheetColor = Partial<ExcelJS.Color> & { indexed?: number; tint?: number };

/** The colour's six hex digits: an explicit ARGB, else an indexed, else a theme colour. */
function baseHex(value: SpreadsheetColor, theme: readonly string[]): string | undefined {
  const indexed = value.indexed === undefined ? undefined : palette[value.indexed],
    themed = value.theme === undefined ? undefined : theme[value.theme];
  return value.argb?.slice(-6) ?? indexed ?? themed;
}

/** Hue in degrees of red, green and blue channels between 0 and 1. */
function hueOf([red, green, blue]: number[], max: number, delta: number): number {
  if (delta === 0) return 0;
  const hue =
    max === red
      ? ((green - blue) / delta) % 6
      : max === green
        ? (blue - red) / delta + 2
        : (red - green) / delta + 4;
  return (hue * 60 + 360) % 360;
}

/** The channels before lightness offset, by which sixth of the hue circle the hue is in. */
function sectorChannels(hue: number, chroma: number, cross: number): number[] {
  const sectors = [
    [chroma, cross, 0],
    [cross, chroma, 0],
    [0, chroma, cross],
    [0, cross, chroma],
    [cross, 0, chroma],
    [chroma, 0, cross],
  ];
  return sectors[Math.min(5, Math.floor(hue / 60))];
}

/** A hex colour lightened (positive tint) or darkened (negative tint) in HSL, as Excel does. */
function tinted(hex: string, tintValue: number): string {
  const channels = [0, 2, 4].map((index) => parseInt(hex.slice(index, index + 2), 16) / 255);
  const max = Math.max(...channels),
    min = Math.min(...channels),
    delta = max - min;
  const lightness = (max + min) / 2;
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  const hue = hueOf(channels, max, delta);
  const tint = Math.max(-1, Math.min(1, tintValue));
  const light = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const chroma = (1 - Math.abs(2 * light - 1)) * saturation,
    cross = chroma * (1 - Math.abs(((hue / 60) % 2) - 1)),
    offset = light - chroma / 2;
  const digits = sectorChannels(hue, chroma, cross).map((value) =>
    Math.round((value + offset) * 255)
      .toString(16)
      .padStart(2, '0'),
  );
  return '#' + digits.join('');
}

export function excelColor(
  value: SpreadsheetColor | undefined,
  theme: readonly string[] = [],
): string | undefined {
  const hex = value && baseHex(value, theme);
  if (!hex || !/^[a-f\d]{6}$/i.test(hex)) return;
  return value.tint ? tinted(hex, value.tint) : '#' + hex.toLowerCase();
}
