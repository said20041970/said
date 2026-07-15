import React from "react";
import { AbsoluteFill, Sequence } from "remotion";
import { z } from "zod";
import "./fonts";
import { GoldParticles } from "./GoldParticles";
import { LogoIntro, MiniLogo } from "./Logo";
import { Stat } from "./Stat";
import { CTA } from "./CTA";

export const goldenPromoSchema = z.object({});

const LOGO_INTRO_FRAMES = 40;
const STATS_START = 90; // 3s @ 30fps
const STATS_END = 300; // 10s @ 30fps
const STAT_COUNT = 3;
const STAT_DURATION = (STATS_END - STATS_START) / STAT_COUNT;

export const GoldenPromo: React.FC = () => {
  return (
    <AbsoluteFill style={{ backgroundColor: "#000" }}>
      <GoldParticles />

      <Sequence durationInFrames={LOGO_INTRO_FRAMES}>
        <LogoIntro />
      </Sequence>

      <Sequence from={LOGO_INTRO_FRAMES - 10}>
        <MiniLogo />
      </Sequence>

      <Sequence from={STATS_START} durationInFrames={STAT_DURATION}>
        <Stat value={500} suffix="+" label="عميل سعيد" />
      </Sequence>
      <Sequence from={STATS_START + STAT_DURATION} durationInFrames={STAT_DURATION}>
        <Stat value={98} suffix="%" label="نسبة الرضا" />
      </Sequence>
      <Sequence from={STATS_START + STAT_DURATION * 2} durationInFrames={STAT_DURATION}>
        <Stat value={24} prefix="" suffix="/7" label="دعم فني" />
      </Sequence>

      <Sequence from={STATS_END}>
        <CTA />
      </Sequence>
    </AbsoluteFill>
  );
};
