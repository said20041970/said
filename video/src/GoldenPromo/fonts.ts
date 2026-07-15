import { loadFont } from "@remotion/fonts";
import { staticFile } from "remotion";

// Unicode subsets copied from the Cairo font (Google Fonts), so the correct
// glyph set is picked automatically depending on script.
const ARABIC_RANGE =
  "U+0600-06FF,U+0750-077F,U+0870-088E,U+0890-0891,U+0897-08E1,U+08E3-08FF,U+200C-200E,U+2010-2011,U+204F,U+2E41,U+FB50-FDFF,U+FE70-FE74,U+FE76-FEFC";
const LATIN_RANGE =
  "U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD";

export const FONT_FAMILY = "Cairo";

const WEIGHTS = ["400", "700", "800", "900"] as const;

// Registered at module scope (not inside a component) so the delayRender()
// call that loadFont() performs internally is set up before the first
// render pass, guaranteeing glyphs are ready before any frame is captured.
export const fontsLoaded: Promise<void> = Promise.all(
  WEIGHTS.flatMap((weight) => [
    loadFont({
      family: FONT_FAMILY,
      url: staticFile(`fonts/cairo-arabic-${weight}-normal.woff2`),
      weight,
      unicodeRange: ARABIC_RANGE,
    }),
    loadFont({
      family: FONT_FAMILY,
      url: staticFile(`fonts/cairo-latin-${weight}-normal.woff2`),
      weight,
      unicodeRange: LATIN_RANGE,
    }),
  ]),
).then(() => undefined);
