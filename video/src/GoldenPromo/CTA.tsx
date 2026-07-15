import React from "react";
import { AbsoluteFill, interpolate, spring, useCurrentFrame, useVideoConfig } from "remotion";
import { FONT_FAMILY } from "./fonts";

export const CTA: React.FC = () => {
  const frame = useCurrentFrame();
  const { fps } = useVideoConfig();

  const bgFade = interpolate(frame, [0, 25], [0, 1], {
    extrapolateLeft: "clamp",
    extrapolateRight: "clamp",
  });
  const entrance = spring({ frame, fps, durationInFrames: 22, config: { damping: 14, mass: 0.6 } });
  const pulse = 1 + 0.035 * Math.sin(frame * 0.15);

  return (
    <AbsoluteFill>
      <AbsoluteFill
        style={{
          opacity: bgFade,
          background:
            "radial-gradient(circle at 50% 45%, rgba(255, 205, 90, 0.35) 0%, rgba(120, 80, 10, 0.18) 45%, rgba(0,0,0,0) 75%)",
        }}
      />
      <AbsoluteFill style={{ justifyContent: "center", alignItems: "center" }}>
        <div
          style={{
            opacity: entrance,
            transform: `scale(${(0.8 + entrance * 0.2) * pulse})`,
            padding: "30px 88px",
            borderRadius: 70,
            background: "linear-gradient(135deg, #FFE27A 0%, #FFD700 40%, #B8860B 100%)",
            boxShadow: "0 0 70px rgba(255, 215, 0, 0.65)",
          }}
        >
          <span
            style={{
              fontFamily: FONT_FAMILY,
              fontWeight: 900,
              fontSize: 64,
              color: "#1a1204",
            }}
          >
            ابدأ الآن
          </span>
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
};
