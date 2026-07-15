import React from "react";
import { AbsoluteFill, interpolate, spring, useCurrentFrame, useVideoConfig } from "remotion";
import { FONT_FAMILY } from "./fonts";

const BRAND_NAME = "FANTAISIE";

export const LogoIntro: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps, durationInFrames } = useVideoConfig();

  const entrance = spring({ frame, fps, durationInFrames: 24, config: { damping: 14, mass: 0.6 } });
  const exit = interpolate(frame, [durationInFrames - 12, durationInFrames], [1, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  const glow = 0.65 + 0.35 * Math.sin(frame * 0.25);
  const scale = 0.7 + entrance * 0.3;

  return (
    <AbsoluteFill style={{ justifyContent: "center", alignItems: "center" }}>
      <div
        style={{
          fontFamily: FONT_FAMILY,
          fontWeight: 900,
          fontSize: 120,
          letterSpacing: 14,
          color: "#FFE27A",
          opacity: entrance * exit,
          transform: `scale(${scale})`,
          textShadow: `0 0 ${30 + glow * 40}px rgba(255, 214, 90, ${0.55 + glow * 0.35}), 0 0 ${
            70 + glow * 60
          }px rgba(255, 180, 40, ${0.35 + glow * 0.25})`,
          textAlign: "center",
        }}
      >
        {BRAND_NAME}
      </div>
    </AbsoluteFill>
  );
};

export const MiniLogo: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();
  const entrance = spring({ frame, fps, durationInFrames: 20, config: { damping: 16 } });

  return (
    <AbsoluteFill style={{ justifyContent: "flex-start", alignItems: "center", paddingTop: 90 }}>
      <div
        style={{
          fontFamily: FONT_FAMILY,
          fontWeight: 800,
          fontSize: 34,
          letterSpacing: 8,
          color: "#FFD966",
          opacity: entrance * 0.9,
          textShadow: "0 0 18px rgba(255, 214, 90, 0.45)",
        }}
      >
        {BRAND_NAME}
      </div>
    </AbsoluteFill>
  );
};
