import React from "react";
import { AbsoluteFill, interpolate, spring, useCurrentFrame, useVideoConfig } from "remotion";
import { FONT_FAMILY } from "./fonts";

interface StatProps {
  value: number;
  prefix?: string;
  suffix?: string;
  label: string;
}

export const Stat: React.FC<StatProps> = ({ value, prefix = "", suffix = "", label }) => {
  const frame = useCurrentFrame();
  const { fps, durationInFrames } = useVideoConfig();

  const entrance = spring({ frame, fps, durationInFrames: 18, config: { damping: 15, mass: 0.6 } });
  const countProgress = spring({ frame, fps, durationInFrames: 28, config: { damping: 200 } });
  const exit = interpolate(frame, [durationInFrames - 10, durationInFrames], [1, 0], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });

  const count = Math.round(countProgress * value);

  return (
    <AbsoluteFill style={{ justifyContent: "center", alignItems: "center" }}>
      <div
        style={{
          direction: "rtl",
          textAlign: "center",
          opacity: entrance * exit,
          transform: `scale(${0.85 + entrance * 0.15})`,
        }}
      >
        <div
          style={{
            fontFamily: FONT_FAMILY,
            fontWeight: 900,
            fontSize: 150,
            color: "#FFD700",
            textShadow: "0 0 40px rgba(255, 215, 0, 0.55)",
            lineHeight: 1,
          }}
        >
          {prefix}
          {count}
          {suffix}
        </div>
        <div
          style={{
            fontFamily: FONT_FAMILY,
            fontWeight: 600,
            fontSize: 46,
            color: "#FFFFFF",
            marginTop: 26,
          }}
        >
          {label}
        </div>
      </div>
    </AbsoluteFill>
  );
};
