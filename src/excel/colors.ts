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
export function excelColor(
  value: (Partial<ExcelJS.Color> & { indexed?: number; tint?: number }) | undefined,
  theme: readonly string[] = [],
): string | undefined {
  const hex =
    value?.argb?.slice(-6) ??
    (value?.indexed === undefined ? undefined : palette[value.indexed]) ??
    (value?.theme === undefined ? undefined : theme[value.theme]);
  if (!hex || !/^[a-f\d]{6}$/i.test(hex)) return;
  if (!value?.tint) return '#' + hex.toLowerCase();
  const channels = [0, 2, 4].map((index) => parseInt(hex.slice(index, index + 2), 16) / 255);
  const max = Math.max(...channels),
    min = Math.min(...channels),
    delta = max - min;
  const lightness = (max + min) / 2;
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  let hue =
    delta === 0
      ? 0
      : max === channels[0]
        ? ((channels[1] - channels[2]) / delta) % 6
        : max === channels[1]
          ? (channels[2] - channels[0]) / delta + 2
          : (channels[0] - channels[1]) / delta + 4;
  hue = (hue * 60 + 360) % 360;
  const tint = Math.max(-1, Math.min(1, value.tint));
  const light = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
  const chroma = (1 - Math.abs(2 * light - 1)) * saturation,
    cross = chroma * (1 - Math.abs(((hue / 60) % 2) - 1)),
    offset = light - chroma / 2;
  const rgb =
    hue < 60
      ? [chroma, cross, 0]
      : hue < 120
        ? [cross, chroma, 0]
        : hue < 180
          ? [0, chroma, cross]
          : hue < 240
            ? [0, cross, chroma]
            : hue < 300
              ? [cross, 0, chroma]
              : [chroma, 0, cross];
  return (
    '#' +
    rgb
      .map((value) =>
        Math.round((value + offset) * 255)
          .toString(16)
          .padStart(2, '0'),
      )
      .join('')
  );
}
