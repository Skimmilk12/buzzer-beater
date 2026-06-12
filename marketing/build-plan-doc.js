const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
        ShadingType, ExternalHyperlink, PageBreak } = require('docx');

const OUT = 'C:/Users/kinsm/OneDrive/Desktop/BUZZER BEATER - Build Report and Growth Plan.docx';

// ── helpers ────────────────────────────────────────────────
const t = (text, opts) => new TextRun(Object.assign({ text }, opts || {}));
const p = (children, opts) => new Paragraph(Object.assign({
  children: Array.isArray(children) ? children : [t(children)],
  spacing: { after: 120 },
}, opts || {}));
const h1 = text => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(text)] });
const h2 = text => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(text)] });
const li = (children, ref) => new Paragraph({
  numbering: { reference: ref || 'bullets', level: 0 },
  spacing: { after: 60 },
  children: Array.isArray(children) ? children : [t(children)],
});
const link = (text, url) => new ExternalHyperlink({
  children: [new TextRun({ text, style: 'Hyperlink' })], link: url,
});

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows, headerShade){
  const total = widths.reduce((a,b)=>a+b,0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((cells, ri) => new TableRow({
      children: cells.map((cell, ci) => new TableCell({
        borders,
        width: { size: widths[ci], type: WidthType.DXA },
        shading: ri === 0 && headerShade !== false
          ? { fill: 'FFE3C2', type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          spacing: { after: 0 },
          children: [t(cell, ri === 0 && headerShade !== false ? { bold: true } : {})],
        })],
      })),
    })),
  });
}

// ── document ───────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: 'B34700' },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: '1F3864' },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: 'steps', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 },
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
    },
    children: [
      // ── title ──
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 60 },
        children: [t('BUZZER BEATER', { bold: true, size: 52, color: 'B34700' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [t('Build Report & Growth Plan', { size: 30, color: '1F3864' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
        children: [t('June 9, 2026  ·  Robert (Skimmilk12) + Claude  ·  v1.2.0 shipped', { italics: true, size: 20, color: '666666' })] }),
      p([t('Purpose of this document: ', { bold: true }),
         t('a briefing for any LLM or collaborator joining the project. It assumes comfort with code and game-design terminology. North star: a simple, addictive, one-button game (the Flappy Bird / Dunk Shot lane) grown into a real income stream.')]),

      // ── 1. what exists ──
      h1('1. What Has Been Built (v1.2.0, live)'),
      p([t('The game: ', { bold: true }),
         t('one-button rooftop basketball. Hold to charge a power meter, release to shoot. A teal "swish band" on the meter guarantees a swish when hit; it shrinks as score climbs. 24-second shot clock; makes add time; hoop relocates per make. Streak multipliers (NBA Jam-style ON FIRE), quarters that change rules every 6 makes (wind, double rim, sprinting meter, half band, flickering meter, then CRUNCH TIME stacks them endlessly), twin-hoop possessions where release power literally chooses between a safe 2 and a deep 3, deterministic HEAT CHECK golden hoops every 4th straight make, and the signature mechanic: a ball airborne when the horn sounds can be banked for triple points and chainable overtime.')]),
      h2('Architecture (for an LLM picking this up)'),
      li([t('Single self-contained HTML file (~2,400 lines): vanilla JS, canvas renderer, WebAudio (every sound and the lo-fi music loop is synthesized — zero asset files), fonts base64-inlined, PWA manifest + network-first service worker. No build step, no dependencies.')]),
      li([t('Dual world geometry: a LAYOUTS table (landscape 960x540 desktop, portrait 540x960 one-handed phone) sharing one physics/logic core. Layout picked by viewport aspect at boot/restart.')]),
      li([t('Design contract #1 — the band promise: ', { bold: true }), t('a release inside the band ALWAYS swishes. Enforced mechanically (snapped balls bypass rim colliders; a numeric solver computes exact launch speed including wind). Verified 40/40 in both orientations under every modifier combination. Never ship anything that breaks this — it is the fairness contract the whole game stands on.')]),
      li([t('Design contract #2 — daily determinism: ', { bold: true }), t('all gameplay randomness flows through a seeded RNG (date-seeded in Daily mode; exactly six draws per possession), so every player on Earth gets the same daily run. Missions are date-seeded the same way.')]),
      li([t('Physics note: near-vertical hoops (portrait) cannot use a capped launch angle — arcs graze the rim at apex. Steep shots construct the arc from a guaranteed 100px apex clearance instead (shotAngle()).')]),
      li([t('Retention systems already in: daily challenge + streak with one auto-shield per week (Duolingo model), 3 rotating date-seeded missions paying career stars, 6 skill badges unlocking ball skins, personal-best chase callouts, near-miss rim rattles in slow-mo, rising-pitch swish combo audio, one reserved mega-celebration for new personal bests, Wordle-style spoiler-free share card with the run story and the game URL.')]),
      li([t('Test methodology: a synchronous in-browser harness drives the real game loop (fireSync/tickGame) — physics regression suites run in seconds and are immune to background-tab throttling. Three multi-agent adversarial reviews caught and fixed 44 real bugs pre-ship.')]),
      h2('Distribution (live today, all $0 cost)'),
      li([t('Web (own URL + PWA): '), link('https://skimmilk12.github.io/buzzer-beater/', 'https://skimmilk12.github.io/buzzer-beater/'), t('  — installable on phones, works offline. Share button carries this URL.')]),
      li([t('itch.io storefront: '), link('skimmilk12.itch.io/buzzer-beater', 'https://skimmilk12.itch.io/buzzer-beater'), t(' — published, free + donations ($2 suggested), comments on, AI disclosure filed. Save export/import protects players from itch iframe storage wipes.')]),
      li([t('Repo: github.com/Skimmilk12/buzzer-beater (public). Deploys: git push (Pages) + butler push (itch) — both one-command.')]),

      // ── 2. money path ──
      new Paragraph({ children: [new PageBreak()] }),
      h1('2. Income Plan & Timeline'),
      p([t('Research-backed thesis (June 2026): ', { bold: true }),
         t('for a no-audience solo dev, web game portals are the only channel that supplies millions of players for $0. itch is a storefront, not a paycheck (most games earn <$100 lifetime). Paid-upfront mobile is dead for unknown devs. The realistic ladder is: portals for revenue + validation, then mobile free-with-ads + IAP, then (content permitting) Steam.')]),
      table([2600, 4200, 2560], [
        ['Channel', 'Realistic monthly revenue', 'Basis'],
        ['CrazyGames (45M players/mo)', '$50 - $500', 'indie dev reports, rev share from launch'],
        ['Poki (100M players/mo)', '$500 - $3,000 (first game)', "Poki's own stated range; top studios reach $1M/yr"],
        ['Non-exclusive licenses', '$500 - $2,000 each (one-time)', 'Coolmath, GameDistribution market rates'],
        ['iOS free + ads + IAP', '$0.02-$0.10 ARPDAU; meaningful only with traction', 'rewarded eCPM $15-40, interstitial $5-15 (US)'],
        ['itch donations', '~$0 (treat as tip jar)', '1-3% conversion on free games'],
      ]),
      p(''),
      h2('Timeline'),
      table([2200, 4600, 2560], [
        ['When', 'Milestone', 'Gate / cost'],
        ['Weeks 1-2', 'Phase 3a: portal SDK adapter (one interface wrapping both CrazyGames + Poki SDKs: gameplayStart/Stop, interstitials between runs, rewarded ads). Submit to CrazyGames.', '$0'],
        ['Weeks 2-6', 'CrazyGames QA -> 2-week Basic Launch -> Full Launch w/ monetization. Submit to Poki (Web Fit Test: ~10K real players, 3-5 days, free professional validation).', 'acceptance = retention'],
        ['Months 2-3', 'Iterate on portal analytics (D1/D7, session length). Build currency + cosmetics catalog (section 3) and 1-2 new modes/themes (section 4). Shop non-exclusive licenses in parallel.', '$0'],
        ['Months 3-4', 'Phase 4: iPhone via Capacitor 8 + Codemagic CI (builds AND signs from Windows - no Mac). TestFlight -> App Store: free + ads + $2.99 remove-ads + coin IAP.', 'gate: portal acceptance OR ~1K plays w/ D1>25%; $99/yr Apple'],
        ['Months 4-5', 'Android via same Capacitor project (Google Play).', '$25 one-time'],
        ['Months 5-6+', 'Steam "Deluxe" consideration: requires real content depth (modes, themes, achievements). Wrap via Electron/Tauri + Steamworks. Price $2.99-4.99.', '$100/app Steam Direct'],
      ]),
      p(''),
      p([t('Honest expectations: ', { bold: true }),
         t('months 1-2 likely earn $0-300 while portal review runs. If Poki accepts, $500-3,000/mo is their stated first-game range and compounds with updates. If both portals decline, the fallback is license sales ($500-2K each) plus iterating on their feedback and resubmitting — rejection feedback from a 10K-player test is itself worth more than most paid playtesting.')]),

      // ── 3. monetization design ──
      new Paragraph({ children: [new PageBreak()] }),
      h1('3. Monetization Design: Currency & Cosmetics'),
      p([t('Platform split that keeps everyone happy: ', { bold: true }),
         t('portal builds monetize via their ad SDKs only (portals forbid external payments) — interstitials between runs, rewarded ads gating bonuses ("watch to continue this run once", "double daily reward"). The owned platforms (iOS/Android/Steam/own site) carry the in-game currency and IAP described below.')]),
      h2('Currency: Swish Coins'),
      li([t('Earn free (deliberately slow): ', { bold: true }), t('~1 coin per 25 points scored, +10 per completed mission star, +5 daily-streak bonus per consecutive day, +2 per golden hoop cashed. A casual day of play earns ~15-30 coins; an average cosmetic costs ~300 — about 1.5-2 weeks of free play per item. Paid is instant. This is the classic fair model: money buys time, never advantage.')]),
      li([t('Buy (IAP packs, iOS/Android/Steam): ', { bold: true }), t('$1.99 = 500 - $4.99 = 1,500 - $9.99 = 3,500. Plus a one-time $2.99 REMOVE ADS. Later: a $4.99 cosmetic-only "Season Ticket" pass per 8-week season.')]),
      h2('Cosmetic catalog (all visual, zero gameplay effect)'),
      table([3400, 1800, 4160], [
        ['Item class', 'Coin price', 'Notes'],
        ['Basketballs (skins + animated rares)', '150-800', '7 already exist via badges; expand to 25+'],
        ['Ball trails (incl. prism, fire, frost)', '100-300', 'cheap, high visibility in shares'],
        ['Shooter characters (silhouettes w/ idle anims)', '300-600', 'hoodie kid, old-timer, mascot, robot...'],
        ['Announcer packs (streak callout styles + stings)', '400-800', 'text-style + horn packs first; recorded VO later'],
        ['Court themes / backgrounds (see section 4)', '500-1,000', 'biggest visual purchase, pairs with modes'],
        ['Lettering/font packs (HUD + popup styles)', '200-400', 'neon, chalk, graffiti, retro-LCD'],
        ['Net styles + swish celebration effects', '100-300', 'chains, ribbons, fireworks, confetti styles'],
      ]),
      p(''),
      h2('Hard rules (do not violate)'),
      li('Cosmetic-only forever: nothing purchasable may touch the band, physics, scoring, or clock. The daily must stay a pure skill comparison or the share loop dies.'),
      li('The 6 badge skins stay unbuyable — prestige money cannot buy is what makes the rest worth buying.'),
      li('No energy systems, no loot boxes, no pay-gated play. Timers may gate bonuses, never access.'),

      // ── 4. product roadmap ──
      h1('4. Making the Game Better: Modes, Themes, Obstacles'),
      h2('New modes (each is a small variation on the verified core)'),
      li([t('Time Attack: ', { bold: true }), t('fixed 60 seconds, no clock bonuses — pure rate-of-fire skill, separate leaderboard line in the share card.')]),
      li([t('Weekly Gauntlet: ', { bold: true }), t('7-day fixed seed and modifier order, one attempt per day — appointment play with higher stakes than the daily.')]),
      li([t('Ghost Duels (H-O-R-S-E): ', { bold: true }), t('deterministic engine + recorded inputs = full replays in a tiny share code. Friend taps your link, plays against your ghost shot-for-shot. This is the single highest-leverage social feature on the roadmap.')]),
      li([t('Zen: ', { bold: true }), t('no clock, no score pressure — practice + the place players test new cosmetics.')]),
      h2('Themes / backgrounds (sold as cosmetics, section 3)'),
      li('Midnight neon city - beach sunset - winter rooftop (snow physics cosmetic only) - space station (visual low-gravity flavor; actual gravity unchanged to protect the band) - championship stadium.'),
      h2('Obstacles & modifiers (quarter pool additions)'),
      li('Pigeons crossing the arc lane (cosmetic dodge drama; never block snapped shots), gusting wind that shifts once mid-flight (solver-compensated at release, honest), defender hand near the rim that swats only non-band releases, balcony ledges that create deliberate bank-shot lines.'),
      li([t('Rule for every new modifier: ', { bold: true }), t('the band promise survives it, verified by the 40-shot suite before ship.')]),
      h2('Polish ladder'),
      li('Announcer v1 (styled callouts + crowd stings) -> recorded VO packs; slow-mo replay of the run’s best shot on the game-over screen (auto-generates the shareable/clippable moment for TikTok); haptics on iOS; localization (the game is nearly text-free - cheap to localize).'),

      // ── 5. integrations ──
      h1('5. Platform Integrations'),
      li([t('iPhone (Phase 4): ', { bold: true }), t('Capacitor 8 wraps the HTML file natively; Codemagic free tier (500 macOS min/mo) builds, signs via App Store Connect API key, and uploads to TestFlight entirely from Windows — no Mac. Config: portrait lock, scroll/zoom disabled, haptics plugin, AVAudioSession playback (ignores mute switch like real games), everything bundled in-binary (then Apple judges only on feel, guideline 4.2). Free + ads + IAP under the 15% Small Business Program.')]),
      li([t('Android: ', { bold: true }), t('same Capacitor project, $25 one-time Play fee. Do after iOS validates.')]),
      li([t('Steam: ', { bold: true }), t('viable only as a "Deluxe" build with real content depth (all modes/themes, Steamworks achievements, local high-score seasons). Wrap with Electron or Tauri + Steamworks bindings. $100/app fee, $2.99-4.99 price. Hypercasual sells poorly on Steam, so this waits for content and an audience — but it is the bridge to the original Steam dream and to bigger Unity projects later.')]),
      li([t('Also HTML5-native and worth a look post-portals: ', { bold: true }), t('Discord Activities and YouTube Playables (Poki’s web-exclusive deal explicitly leaves app stores and these surfaces negotiable — read terms before signing).')]),

      // ── 6. growth ──
      h1('6. Growth Playbook'),
      li([t('The share loop is the engine: ', { bold: true }), t('every game-over offers a one-tap, spoiler-free run story + URL. Daily seed makes scores comparable; missions give three extra reasons to open it daily.')]),
      li('Clip channel: the slow-mo buzzer-beater replays are TikTok/Shorts-native 10-second clips. Post 2-3/week; the game title is burned into the HUD.'),
      li('Communities: r/WebGames, r/browsergames, itch devlogs (each update = a devlog post = a discovery event), web-game Discords. Micro-streamers who play browser games accept game links readily.'),
      li('Analytics: portal dashboards give D1/D7/session length (the numbers that gate Phase 4); add a privacy-friendly counter (GoatCounter, no cookie banner) to the Pages build.'),
      li('Cadence: one visible update per week (new mission types, a cosmetic, a modifier) — portals weight actively-updated games, and the devlog/clip pipeline feeds itself.'),

      // ── 7. risks ──
      h1('7. Risks & Mitigations'),
      table([2800, 6560], [
        ['Risk', 'Mitigation'],
        ['Portal rejection', 'Web Fit Test feedback is itself valuable; iterate and resubmit; fallback to license sales ($500-2K each) and CrazyGames-only revenue.'],
        ['Discoverability stall', 'The game is share-architected (daily + story cards + ghost duels next); growth compounds from players, not store browsing.'],
        ['Solo bandwidth', 'AI-assisted development has shipped 3 verified releases in one day; the single-file architecture keeps iteration cheap. Guard the test harness and the two design contracts.'],
        ['AI-content filters (itch) / policy shifts', 'Disclosure already filed honestly; revenue plan does not depend on itch. Portal/app-store policies re-checked at each phase.'],
        ['Scope creep killing the simple lane', 'Every feature must pass: "can a first-timer still understand the game in 10 seconds?" The roguelite was already cut on this rule.'],
      ]),

      // ── 8. next actions ──
      h1('8. Immediate Next Actions'),
      li('Robert plays v1.2.0 portrait on his phone and reports feel (charge timing, thumb comfort, first-session clarity).', 'steps'),
      li('Create CrazyGames developer account; build the portal SDK adapter; integrate + smoke-test with their QA tool; submit.', 'steps'),
      li('Submit to Poki (Web Fit Test) in the same window.', 'steps'),
      li('While reviews run: implement Swish Coins + first cosmetic shelf (owned-platform builds only), Time Attack mode, and the slow-mo best-shot replay (share-clip generator).', 'steps'),
      li('On portal acceptance or 1K plays w/ D1>25%: enroll Apple Developer ($99) and start the Capacitor build.', 'steps'),
      p(''),
      new Paragraph({ spacing: { before: 200 },
        children: [t('Built for Robert - vibe-coded with Claude - June 2026', { italics: true, size: 18, color: '888888' })] }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('written: ' + OUT + ' (' + buf.length + ' bytes)');
});
