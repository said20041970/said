import React, { useMemo } from "react";
import { AbsoluteFill, random, useCurrentFrame, useVideoConfig } from "remotion";

const PARTICLE_COUNT = 55;

type ParticleSpec = {
  x: number;
  startY: number;
  size: number;
  speed: number;
  swayAmplitude: number;
  swaySpeed: number;
  phase: number;
  opacityBase: number;
};

const makeParticles = (count: number, width: number, height: number): ParticleSpec[] => {
  return new Array(count).fill(0).map((_, i) => {
    const seed = `gold-particle-${i}`;
    return {
      x: random(`${seed}-x`) * width,
      startY: random(`${seed}-y`) * height,
      size: 2 + random(`${seed}-size`) * 6,
      speed: 0.3 + random(`${seed}-speed`) * 0.9,
      swayAmplitude: 8 + random(`${seed}-sway`) * 26,
      swaySpeed: 0.015 + random(`${seed}-swaySpeed`) * 0.035,
      phase: random(`${seed}-phase`) * Math.PI * 2,
      opacityBase: 0.25 + random(`${seed}-opacity`) * 0.55,
    };
  });
};

export const GoldParticles: React.FC = () => {
  const frame = useCurrentFrame();
  const { width, height } = useVideoConfig();

  const particles = useMemo(() => makeParticles(PARTICLE_COUNT, width, height), [width, height]);

  return (
    <AbsoluteFill style={{ backgroundColor: "#050302" }}>
      {particles.map((p, i) => {
        const wrap = height + 120;
        const travelled = ((p.startY - frame * p.speed) % wrap + wrap) % wrap;
        const y = travelled - 60;
        const x = p.x + Math.sin(frame * p.swaySpeed + p.phase) * p.swayAmplitude;
        const twinkle = 0.6 + 0.4 * Math.sin(frame * 0.1 + p.phase);

        return (
          <div
            key={i}
            style={{
              position: "absolute",
              left: x,
              top: y,
              width: p.size,
              height: p.size,
              borderRadius: "50%",
              backgroundColor: "#FFD54A",
              boxShadow: `0 0 ${p.size * 4}px ${p.size}px rgba(255, 200, 60, ${0.55 * twinkle})`,
              opacity: p.opacityBase * twinkle,
            }}
          />
        );
      })}
    </AbsoluteFill>
  );
};
