/**
 * The suite's own English thesaurus: sets of words that mean the same thing
 * in one sense, one set a line, each with its part of speech — n noun, v verb,
 * a adjective, r adverb. A word in several sets has several meanings. Written
 * for this suite, in British spelling, for the words a document most often
 * reaches for; thesaurus.js reads it.
 */
export const THESAURUS_EN = `
a: happy, glad, cheerful, joyful, merry, content, contented, pleased, delighted, jolly, upbeat
a: sad, unhappy, sorrowful, downcast, dejected, gloomy, miserable, melancholy, despondent, glum, low, blue
a: angry, cross, furious, irate, mad, livid, annoyed, indignant, enraged, incensed, fuming
a: afraid, frightened, scared, fearful, terrified, alarmed, nervous, anxious, timid
a: brave, courageous, bold, fearless, daring, heroic, valiant, plucky, intrepid, gallant
a: big, large, huge, enormous, vast, immense, great, massive, giant, gigantic, sizeable, substantial
a: small, little, tiny, minute, miniature, compact, petite, minuscule, slight, wee
a: good, fine, excellent, great, superb, splendid, first-rate, admirable, marvellous, wonderful
a: bad, poor, inferior, substandard, awful, terrible, dreadful, abysmal, shoddy, deficient
a: fast, quick, rapid, swift, speedy, brisk, hasty, prompt, nimble, fleet
a: slow, unhurried, leisurely, sluggish, gradual, plodding, dawdling, lagging
a: easy, simple, straightforward, effortless, uncomplicated, painless, elementary, undemanding
a: hard, difficult, tough, demanding, challenging, arduous, laborious, strenuous, taxing
a: hard, firm, solid, rigid, stiff, unyielding, dense, tough
a: soft, tender, gentle, smooth, supple, yielding, pliable, squashy, velvety
a: beautiful, lovely, pretty, attractive, gorgeous, stunning, handsome, charming, elegant, exquisite
a: ugly, unattractive, unsightly, hideous, plain, grotesque, unlovely
a: clever, intelligent, bright, smart, brilliant, gifted, able, quick-witted, sharp, astute, shrewd
a: stupid, foolish, silly, dim, dense, dull-witted, senseless, idiotic, unintelligent
a: rich, wealthy, affluent, prosperous, well-off, well-to-do, moneyed, loaded
a: poor, needy, impoverished, penniless, destitute, hard-up, broke, deprived
a: new, fresh, novel, recent, modern, original, up-to-date, current, latest
a: old, aged, elderly, ancient, antique, vintage, veteran, senior
a: old, former, previous, past, earlier, prior, erstwhile, one-time
a: important, significant, major, vital, essential, crucial, key, central, principal, critical
a: unimportant, minor, trivial, insignificant, petty, negligible, slight, inconsequential
a: true, correct, right, accurate, exact, precise, factual, valid, sound
a: false, untrue, wrong, incorrect, mistaken, inaccurate, erroneous, fallacious
a: honest, truthful, sincere, frank, candid, open, straightforward, upright, trustworthy
a: dishonest, deceitful, untruthful, lying, crooked, fraudulent, shifty, underhand
a: kind, kindly, caring, considerate, thoughtful, generous, compassionate, warm, benevolent, gentle
a: cruel, harsh, unkind, heartless, brutal, callous, ruthless, merciless, vicious
a: calm, peaceful, quiet, still, tranquil, serene, placid, relaxed, composed, untroubled
a: busy, occupied, engaged, active, hectic, bustling, working, tied up
a: idle, inactive, unoccupied, unemployed, lazy, inert, dormant
a: lazy, idle, indolent, slothful, sluggish, workshy, lethargic
a: tired, weary, exhausted, fatigued, worn out, drained, sleepy, drowsy, spent
a: strong, powerful, mighty, sturdy, robust, tough, muscular, forceful, potent
a: weak, feeble, frail, delicate, fragile, flimsy, puny, faint, powerless
a: clean, spotless, immaculate, pristine, washed, hygienic, unsoiled, tidy
a: dirty, filthy, grimy, grubby, soiled, unclean, muddy, dusty, mucky
a: tidy, neat, orderly, trim, well-kept, shipshape, organised
a: messy, untidy, disorderly, chaotic, cluttered, jumbled, scruffy, sloppy
a: hot, warm, boiling, scorching, baking, sweltering, roasting, heated
a: cold, chilly, cool, freezing, frosty, icy, wintry, bitter, nippy
a: wet, damp, moist, soaked, sodden, drenched, soggy, saturated, watery
a: dry, arid, parched, dehydrated, bone-dry, desiccated, thirsty
a: dark, dim, gloomy, murky, shadowy, unlit, sombre, black
a: bright, light, brilliant, dazzling, radiant, vivid, shining, luminous, gleaming
a: loud, noisy, deafening, booming, thunderous, blaring, rowdy, raucous
a: quiet, silent, hushed, soundless, muted, soft, low, noiseless
a: funny, amusing, humorous, comic, comical, hilarious, witty, entertaining, droll
a: serious, solemn, grave, earnest, sober, stern, sombre, thoughtful
a: strange, odd, peculiar, weird, unusual, curious, bizarre, unfamiliar, eerie, queer
a: normal, ordinary, usual, typical, regular, standard, common, routine, conventional
a: common, frequent, widespread, everyday, familiar, prevalent, general, ordinary
a: rare, unusual, uncommon, scarce, infrequent, exceptional, sparse, few
a: certain, sure, confident, positive, definite, convinced, assured, satisfied
a: uncertain, unsure, doubtful, dubious, hesitant, undecided, unclear, vague
a: clear, plain, obvious, evident, apparent, distinct, explicit, transparent, unmistakable
a: vague, unclear, imprecise, hazy, fuzzy, indistinct, ambiguous, woolly
a: complete, whole, entire, full, total, intact, finished, comprehensive
a: partial, incomplete, limited, fragmentary, unfinished, imperfect
a: empty, vacant, bare, hollow, unoccupied, void, blank, deserted
a: full, packed, crammed, crowded, stuffed, loaded, brimming, overflowing
a: safe, secure, protected, sheltered, unharmed, sound, out of danger
a: dangerous, risky, hazardous, perilous, unsafe, treacherous, precarious
a: famous, well-known, renowned, celebrated, noted, eminent, distinguished, prominent, illustrious
a: unknown, obscure, unfamiliar, anonymous, unnamed, little-known, unsung
a: polite, courteous, civil, respectful, well-mannered, gracious, considerate
a: rude, impolite, discourteous, insolent, impertinent, cheeky, offensive, ill-mannered
a: friendly, amiable, affable, genial, sociable, cordial, warm, welcoming, approachable
a: hostile, unfriendly, antagonistic, aggressive, belligerent, cold, unwelcoming
a: proud, pleased, gratified, satisfied, honoured, dignified
a: arrogant, conceited, vain, proud, haughty, smug, boastful, pompous, superior
a: modest, humble, unassuming, unpretentious, shy, reserved, self-effacing
a: shy, timid, bashful, reserved, retiring, diffident, nervous, coy
a: generous, giving, liberal, charitable, lavish, open-handed, unselfish, bountiful
a: mean, stingy, miserly, tight-fisted, parsimonious, penny-pinching, ungenerous
a: cheap, inexpensive, low-cost, economical, affordable, reasonable, budget, cut-price
a: expensive, costly, dear, pricey, high-priced, extravagant, exorbitant, steep
a: early, premature, advance, initial, first, prompt, timely
a: late, overdue, delayed, tardy, belated, behind
a: recent, latest, new, fresh, modern, current, contemporary
a: ancient, old, antique, archaic, prehistoric, primeval, age-old
a: modern, contemporary, current, up-to-date, present-day, recent, new, latest
a: wide, broad, extensive, expansive, spacious, vast, sweeping
a: narrow, thin, slim, slender, tight, cramped, restricted
a: thin, slim, slender, lean, skinny, slight, bony, lanky, svelte
a: fat, overweight, plump, stout, chubby, obese, portly, heavy, tubby
a: tall, high, lofty, towering, soaring, giant, lanky
a: short, little, small, low, squat, diminutive, petite
a: short, brief, concise, succinct, terse, quick, momentary, fleeting
a: long, lengthy, extended, prolonged, protracted, endless, interminable
a: deep, profound, bottomless, yawning, cavernous
a: shallow, superficial, surface, slight, trivial
a: heavy, weighty, hefty, massive, bulky, cumbersome, ponderous
a: light, lightweight, weightless, feathery, flimsy, airy, portable
a: rough, coarse, uneven, bumpy, rugged, jagged, harsh, bristly
a: smooth, even, flat, level, sleek, polished, silky, glossy
a: sharp, keen, pointed, razor-sharp, honed, spiky, cutting
a: blunt, dull, unsharpened, rounded, edgeless
a: blunt, frank, direct, outspoken, plain-spoken, forthright, candid, brusque
a: interesting, fascinating, absorbing, engaging, gripping, intriguing, compelling, riveting
a: boring, dull, tedious, monotonous, uninteresting, tiresome, dreary, humdrum, flat
a: exciting, thrilling, exhilarating, stirring, electrifying, rousing, dramatic
a: pleasant, nice, agreeable, enjoyable, delightful, pleasing, lovely, fine
a: unpleasant, disagreeable, nasty, horrible, objectionable, distasteful, offensive
a: nice, pleasant, agreeable, likeable, kind, lovely, charming, delightful
a: nasty, horrible, unpleasant, vile, foul, disgusting, revolting, repulsive
a: delicious, tasty, appetising, mouth-watering, scrumptious, flavoursome, delectable
a: healthy, well, fit, sound, strong, robust, hale, hearty, wholesome
a: ill, sick, unwell, poorly, ailing, infirm, indisposed, off colour
a: correct, right, accurate, exact, precise, true, proper, faultless
a: wrong, incorrect, mistaken, inaccurate, erroneous, false, faulty
a: possible, feasible, practicable, achievable, attainable, viable, workable, conceivable
a: impossible, unfeasible, unachievable, unworkable, hopeless, inconceivable, out of the question
a: necessary, essential, needed, required, vital, indispensable, requisite, compulsory
a: unnecessary, needless, inessential, superfluous, redundant, uncalled-for, gratuitous
a: available, free, obtainable, accessible, handy, on hand, to hand, ready
a: ready, prepared, set, waiting, organised, primed
a: free, unrestricted, independent, liberated, unconfined, at liberty
a: free, complimentary, gratis, unpaid, without charge, on the house
a: busy, crowded, packed, bustling, lively, hectic, teeming
a: lively, energetic, active, vigorous, spirited, animated, sprightly, dynamic, vivacious
a: dull, boring, flat, lifeless, lacklustre, drab, uninspiring
a: dull, cloudy, overcast, grey, murky, sunless, gloomy
a: simple, plain, basic, unadorned, austere, modest, unfussy
a: complex, complicated, intricate, involved, elaborate, convoluted, tangled
a: obvious, clear, plain, evident, apparent, patent, manifest, conspicuous
a: hidden, concealed, secret, covert, invisible, unseen, disguised, buried
a: secret, confidential, private, classified, hidden, covert, clandestine, undercover
a: public, open, general, communal, shared, common, civic
a: private, personal, confidential, intimate, own, individual, exclusive
a: whole, entire, complete, full, total, undivided, unbroken
a: main, chief, principal, primary, leading, major, central, prime, foremost
a: final, last, closing, concluding, ultimate, terminal, eventual
a: first, initial, opening, earliest, original, primary, introductory
a: next, following, subsequent, succeeding, coming, ensuing, later
a: previous, prior, earlier, preceding, former, past, foregoing
a: various, different, diverse, assorted, varied, miscellaneous, sundry
a: different, dissimilar, unlike, distinct, contrasting, varied, divergent
a: similar, alike, like, comparable, akin, equivalent, matching, analogous
a: same, identical, equal, matching, equivalent, uniform, indistinguishable
a: special, particular, distinctive, exceptional, unique, specific, individual
a: general, broad, overall, universal, widespread, common, comprehensive, sweeping
a: specific, particular, precise, exact, definite, explicit, individual
a: real, genuine, authentic, actual, true, bona fide, valid, legitimate
a: fake, false, counterfeit, imitation, sham, bogus, phoney, forged, artificial
a: natural, normal, unaffected, spontaneous, genuine, instinctive
a: artificial, synthetic, man-made, manufactured, imitation, fake, simulated
a: ripe, mature, ready, mellow, developed, seasoned
a: young, youthful, juvenile, adolescent, junior, immature, infant
a: mature, adult, grown-up, developed, sensible, responsible
a: wise, sensible, sage, prudent, judicious, shrewd, knowledgeable, sound
a: sensible, reasonable, practical, rational, sound, wise, level-headed, logical
a: crazy, mad, insane, deranged, demented, lunatic, unhinged, absurd
a: absurd, ridiculous, ludicrous, preposterous, nonsensical, silly, laughable, farcical
a: patient, tolerant, forbearing, uncomplaining, calm, understanding, long-suffering
a: impatient, restless, eager, edgy, irritable, intolerant, hasty
a: eager, keen, enthusiastic, avid, zealous, willing, impatient, ardent
a: willing, ready, prepared, inclined, happy, glad, disposed, game
a: reluctant, unwilling, hesitant, disinclined, loath, averse, grudging
a: careful, cautious, prudent, wary, watchful, attentive, meticulous, thorough
a: careless, negligent, sloppy, slapdash, thoughtless, reckless, heedless, slipshod
a: thorough, careful, meticulous, painstaking, rigorous, exhaustive, detailed, comprehensive
a: lucky, fortunate, blessed, charmed, favoured, jammy
a: unlucky, unfortunate, hapless, luckless, ill-fated, cursed
a: successful, prosperous, thriving, flourishing, booming, victorious, triumphant
a: useful, helpful, handy, valuable, practical, beneficial, worthwhile, constructive
a: useless, worthless, futile, pointless, ineffective, vain, unhelpful, hopeless
a: valuable, precious, prized, treasured, costly, invaluable, priceless, worthwhile
a: famous, notable, noteworthy, remarkable, celebrated, legendary
a: excellent, outstanding, exceptional, superb, magnificent, first-class, superlative, terrific
a: terrible, awful, dreadful, horrible, appalling, frightful, atrocious, horrendous, dire
a: awful, terrible, dreadful, horrific, horrendous, ghastly, abominable
a: wonderful, marvellous, magnificent, glorious, splendid, superb, fantastic, fabulous, amazing
a: amazing, astonishing, astounding, staggering, stunning, incredible, remarkable, extraordinary
a: surprising, unexpected, startling, unforeseen, astonishing, sudden
a: sudden, abrupt, unexpected, swift, rapid, unforeseen, hasty
a: gradual, slow, steady, gentle, progressive, measured, step-by-step
a: steady, stable, constant, firm, even, regular, unwavering, consistent
a: constant, continual, continuous, perpetual, endless, incessant, unceasing, ceaseless
a: regular, routine, normal, usual, habitual, customary, set, frequent
a: frequent, regular, repeated, recurrent, common, persistent, numerous
a: occasional, intermittent, sporadic, infrequent, irregular, periodic, odd
a: quick, fast, rapid, swift, prompt, speedy, express, immediate
a: immediate, instant, prompt, direct, urgent, swift, at once
a: urgent, pressing, critical, crucial, vital, acute, serious, top-priority
a: grateful, thankful, appreciative, obliged, indebted, beholden
a: sorry, apologetic, regretful, remorseful, contrite, penitent, repentant, ashamed
a: guilty, culpable, responsible, blameworthy, ashamed, at fault
a: innocent, guiltless, blameless, faultless, irreproachable, clear
a: innocent, naive, unworldly, artless, ingenuous, simple, green
a: lonely, alone, lonesome, isolated, solitary, friendless, forlorn
a: alone, solitary, single, lone, isolated, unaccompanied, by oneself
a: famous, popular, fashionable, in vogue, trendy, sought-after
a: popular, well-liked, favourite, admired, in demand, fashionable, accepted
a: fashionable, stylish, chic, trendy, smart, elegant, modern, in vogue
a: ordinary, plain, average, everyday, unremarkable, standard, run-of-the-mill
a: average, typical, normal, ordinary, middling, standard, medium, mean
a: perfect, flawless, faultless, ideal, impeccable, immaculate, exemplary, consummate
a: broken, damaged, smashed, shattered, cracked, faulty, defective, bust
a: whole, unbroken, intact, undamaged, sound, unharmed, perfect
a: open, unlocked, ajar, unfastened, wide open, gaping
a: closed, shut, locked, fastened, sealed, bolted
a: open, frank, candid, honest, direct, transparent, communicative
a: ready, finished, complete, done, prepared, arranged
a: fair, just, impartial, unbiased, even-handed, equitable, objective, neutral
a: unfair, unjust, biased, prejudiced, one-sided, partial, inequitable
a: fair, blond, blonde, light, pale, flaxen, golden
a: pale, white, pallid, ashen, wan, colourless, pasty, washed-out
a: colourful, bright, vivid, vibrant, rich, multicoloured, gaudy, brilliant
a: loyal, faithful, devoted, true, steadfast, staunch, dependable, reliable
a: reliable, dependable, trustworthy, faithful, consistent, steady, sound, responsible
a: famous, historic, momentous, significant, landmark, epoch-making
a: odd, uneven, unmatched, unpaired, single, spare, lone
a: extra, additional, further, more, supplementary, spare, added, surplus
a: spare, extra, surplus, reserve, additional, leftover, free
a: enough, sufficient, adequate, ample, plenty, abundant, satisfactory
a: plentiful, abundant, ample, copious, profuse, lavish, bountiful, rich
a: scarce, rare, sparse, scanty, meagre, insufficient, short, limited
v: say, state, declare, announce, remark, mention, comment, utter, express, voice
v: tell, inform, notify, advise, let know, apprise, report to
v: ask, enquire, inquire, question, query, request, demand, quiz
v: answer, reply, respond, retort, rejoin, acknowledge
v: speak, talk, converse, chat, discuss, communicate, address
v: shout, yell, cry, scream, call, bellow, roar, holler
v: whisper, murmur, mutter, mumble, breathe, hiss
v: walk, stroll, wander, amble, march, stride, saunter, tread, step, pace
v: run, sprint, dash, race, jog, hurry, rush, bolt, gallop, scurry
v: go, leave, depart, set off, set out, exit, withdraw, move
v: come, arrive, approach, appear, reach, turn up, show up
v: look, see, watch, observe, view, regard, eye, glance, gaze, stare
v: see, notice, observe, spot, perceive, detect, discern, glimpse, catch sight of
v: look for, search, seek, hunt, look around, scour, forage
v: find, discover, locate, come across, uncover, unearth, track down, detect
v: lose, mislay, misplace, drop, forget
v: get, obtain, acquire, gain, receive, procure, secure, come by, earn
v: give, hand, present, provide, offer, supply, donate, grant, award, contribute
v: take, grab, seize, grasp, clutch, snatch, catch, pick up
v: take, carry, bring, convey, transport, move, fetch, bear
v: keep, retain, hold, preserve, save, conserve, maintain, store
v: put, place, set, lay, position, deposit, rest, stand
v: make, create, build, produce, form, construct, manufacture, assemble, fashion
v: build, construct, erect, assemble, put up, raise, develop
v: break, smash, shatter, crack, fracture, split, snap, burst
v: mend, repair, fix, restore, patch, renovate, service, put right
v: fix, attach, fasten, secure, fit, join, stick, connect
v: join, connect, link, unite, combine, attach, couple, merge
v: separate, divide, split, part, detach, disconnect, sever, isolate
v: start, begin, commence, initiate, launch, open, set about, embark on
v: stop, halt, end, finish, cease, discontinue, quit, terminate, conclude
v: end, finish, conclude, close, complete, terminate, wind up, finalise
v: continue, carry on, go on, proceed, persist, keep on, resume, maintain
v: change, alter, modify, adjust, vary, transform, convert, amend, revise
v: increase, grow, rise, expand, enlarge, extend, multiply, swell, escalate, boost
v: decrease, reduce, lessen, diminish, decline, drop, fall, shrink, lower, cut
v: grow, develop, expand, increase, mature, sprout, flourish, thrive
v: improve, better, enhance, upgrade, refine, advance, progress, perfect
v: worsen, deteriorate, decline, degenerate, aggravate, slip
v: help, aid, assist, support, back, serve, abet, lend a hand
v: hinder, hamper, obstruct, impede, block, restrict, delay, handicap
v: try, attempt, endeavour, strive, aim, seek, venture, have a go
v: succeed, triumph, win, prevail, prosper, achieve, accomplish, make it
v: fail, flop, fall through, miscarry, founder, collapse, fall short
v: win, triumph, succeed, prevail, conquer, beat, overcome, gain
v: lose, forfeit, be defeated, fail, succumb
v: like, enjoy, love, adore, fancy, appreciate, relish, be fond of
v: love, adore, cherish, treasure, worship, idolise, care for
v: hate, detest, loathe, despise, dislike, abhor, resent
v: want, desire, wish, crave, long for, yearn for, fancy, covet
v: need, require, want, lack, demand, call for
v: think, believe, consider, reckon, suppose, imagine, judge, deem, feel
v: think, ponder, consider, reflect, contemplate, deliberate, muse, mull over
v: know, understand, realise, recognise, comprehend, grasp, appreciate
v: understand, comprehend, grasp, follow, see, fathom, take in, make out
v: remember, recall, recollect, remind, think back, reminisce
v: forget, overlook, neglect, omit, disregard, ignore
v: decide, choose, determine, resolve, settle, conclude, elect, opt
v: choose, select, pick, opt for, elect, decide on, prefer
v: agree, concur, consent, accept, approve, assent, comply, go along with
v: disagree, differ, dissent, object, argue, quarrel, oppose
v: argue, quarrel, disagree, row, bicker, squabble, dispute, debate
v: allow, permit, let, enable, authorise, sanction, tolerate, approve
v: forbid, prohibit, ban, bar, outlaw, prevent, veto, disallow
v: prevent, stop, avert, block, hinder, thwart, preclude, forestall
v: show, display, exhibit, present, reveal, demonstrate, indicate, expose
v: hide, conceal, cover, disguise, mask, bury, screen, veil, secrete
v: explain, clarify, describe, illustrate, interpret, define, spell out, account for
v: describe, portray, depict, report, relate, recount, narrate, characterise
v: suggest, propose, recommend, advise, advocate, put forward, hint, imply
v: offer, propose, tender, present, volunteer, extend, put forward
v: refuse, decline, reject, deny, turn down, rebuff, spurn
v: accept, take, receive, admit, welcome, embrace, acknowledge
v: expect, anticipate, await, foresee, predict, envisage, hope for
v: hope, wish, trust, expect, aspire, long, desire
v: worry, fret, agonise, brood, be anxious, stew
v: frighten, scare, alarm, terrify, startle, intimidate, petrify, panic
v: surprise, astonish, amaze, astound, startle, stun, shock, stagger
v: please, delight, gratify, satisfy, charm, cheer, gladden, content
v: annoy, irritate, anger, bother, exasperate, provoke, vex, irk, nettle, rile
v: bother, disturb, trouble, pester, harass, interrupt, hassle, badger
v: calm, soothe, quieten, settle, pacify, placate, relax, still
v: laugh, giggle, chuckle, chortle, snigger, guffaw, titter, cackle
v: cry, weep, sob, wail, howl, whimper, snivel, blubber
v: smile, grin, beam, smirk, simper
v: eat, consume, devour, swallow, munch, dine, feed, gobble, scoff
v: drink, sip, gulp, swallow, swig, quaff, guzzle, slurp
v: sleep, doze, nap, slumber, snooze, rest, drowse, kip
v: rest, relax, unwind, repose, take it easy, lounge, laze
v: work, labour, toil, slog, operate, function, strive, graft
v: play, amuse oneself, have fun, frolic, romp, sport, game
v: buy, purchase, acquire, obtain, pay for, invest in, procure
v: sell, trade, vend, market, retail, deal in, auction, peddle
v: pay, settle, remit, reimburse, refund, spend, compensate
v: spend, expend, pay out, lay out, use up, consume, squander
v: save, keep, conserve, preserve, store, put aside, hoard, economise
v: save, rescue, free, deliver, liberate, recover, salvage
v: protect, guard, defend, shield, safeguard, shelter, preserve, secure
v: attack, assault, charge, strike, raid, storm, invade, ambush
v: fight, battle, struggle, combat, clash, brawl, scuffle, wrestle
v: hit, strike, beat, punch, slap, smack, thump, knock, bash, whack
v: hurt, injure, harm, wound, damage, maim, bruise, ache
v: kill, slay, murder, slaughter, execute, assassinate, destroy
v: destroy, ruin, wreck, demolish, devastate, wipe out, annihilate, raze
v: move, shift, transfer, relocate, budge, stir, go, transport
v: push, shove, press, thrust, propel, drive, nudge, ram
v: pull, drag, draw, haul, tug, tow, heave, yank
v: lift, raise, elevate, hoist, pick up, heave, uplift
v: drop, fall, descend, plunge, sink, tumble, dip, slump
v: fall, drop, tumble, topple, collapse, plunge, trip, stumble
v: climb, ascend, scale, mount, clamber, scramble, rise
v: jump, leap, spring, bound, hop, vault, skip, pounce
v: throw, toss, fling, hurl, pitch, cast, chuck, lob
v: catch, grab, seize, grasp, capture, trap, snare, intercept
v: hold, grip, grasp, clutch, clasp, carry, cling to, hang on to
v: open, unlock, unfasten, undo, unseal, uncover, unwrap
v: close, shut, lock, fasten, seal, secure, bolt
v: cover, coat, wrap, envelop, cloak, shroud, blanket, overlay
v: fill, pack, load, stuff, cram, top up, stock, refill
v: empty, clear, drain, evacuate, vacate, unload, void
v: clean, wash, cleanse, scrub, wipe, rinse, mop, dust, launder
v: cut, slice, chop, trim, carve, snip, clip, sever, slit
v: write, record, note, jot down, compose, draft, pen, scribble, inscribe
v: read, peruse, scan, study, skim, browse, pore over
v: learn, study, master, pick up, absorb, grasp, discover, find out
v: teach, educate, instruct, train, coach, tutor, school, lecture
v: lead, guide, direct, head, manage, command, steer, conduct, escort
v: follow, pursue, chase, track, trail, shadow, tail, succeed
v: meet, encounter, come across, run into, gather, assemble, convene
v: visit, call on, drop in, see, stop by, look up, tour
v: wait, stay, remain, linger, pause, hang on, hold on, stay put
v: stay, remain, wait, linger, stop, lodge, dwell, reside
v: live, dwell, reside, inhabit, occupy, lodge, stay, settle
v: die, perish, pass away, expire, decease, pass on
v: use, utilise, employ, apply, exploit, operate, exercise, wield
v: need, depend on, rely on, count on, bank on
v: send, dispatch, post, mail, forward, transmit, ship, convey
v: receive, get, obtain, collect, accept, acquire, be given
v: call, phone, ring, telephone, contact, dial
v: call, name, title, dub, label, term, christen, designate
v: invite, ask, request, summon, bid, welcome
v: arrange, organise, plan, order, sort, set up, coordinate, schedule
v: plan, design, devise, draft, map out, scheme, plot, intend
v: prepare, get ready, arrange, organise, make ready, set up, prime
v: check, examine, inspect, test, verify, confirm, review, scrutinise
v: test, try, trial, examine, assess, check, evaluate, experiment
v: prove, demonstrate, show, confirm, establish, verify, substantiate
v: measure, gauge, quantify, calculate, assess, estimate, size up
v: count, total, add up, number, tally, calculate, reckon
v: compare, contrast, liken, match, weigh, juxtapose, equate
v: include, contain, comprise, incorporate, involve, cover, embrace
v: exclude, omit, leave out, rule out, bar, except, eliminate
v: add, append, attach, include, insert, tack on, supplement
v: remove, take away, delete, eliminate, extract, withdraw, erase, strip
v: replace, substitute, swap, exchange, supersede, supplant, succeed
v: return, come back, go back, revert, reappear, recur
v: return, give back, send back, restore, repay, bring back
v: happen, occur, take place, arise, come about, transpire, befall
v: cause, produce, create, lead to, bring about, generate, provoke, trigger
v: affect, influence, impact, alter, change, touch, sway
v: depend, rely, hinge, rest, turn on, hang on
v: seem, appear, look, sound, come across as
v: become, grow, turn, get, come to be, develop into
v: belong, fit, go, be part of, relate
v: own, possess, have, hold, keep, retain
v: lack, need, want, be without, miss, be short of
v: miss, long for, pine for, yearn for, regret
v: miss, skip, omit, overlook, pass over, neglect
v: complain, grumble, moan, protest, object, whine, gripe, carp
v: praise, commend, compliment, applaud, acclaim, laud, admire, extol
v: criticise, condemn, censure, attack, fault, slate, knock, disparage
v: blame, accuse, charge, hold responsible, censure, condemn
v: forgive, pardon, excuse, overlook, absolve, let off
v: thank, acknowledge, credit, appreciate, be grateful
v: apologise, say sorry, express regret, beg pardon
v: promise, pledge, vow, swear, guarantee, undertake, assure
v: warn, caution, alert, advise, notify, forewarn
v: threaten, menace, intimidate, endanger, bully, terrorise
v: encourage, inspire, motivate, urge, cheer, hearten, support, spur
v: discourage, deter, dissuade, dishearten, put off, daunt, demoralise
v: persuade, convince, coax, induce, talk into, win over, sway, influence
v: force, compel, oblige, make, coerce, pressure, require, drive
v: obey, comply, follow, observe, heed, abide by, submit
v: imagine, picture, envisage, visualise, fancy, conceive, dream up
v: invent, create, devise, design, originate, conceive, think up, concoct
v: copy, imitate, duplicate, reproduce, replicate, mimic, mirror, echo
v: develop, evolve, grow, progress, advance, mature, expand, unfold
v: produce, make, manufacture, create, generate, yield, output
v: manage, run, direct, control, administer, supervise, oversee, handle, head
v: control, manage, regulate, govern, direct, command, rule, restrain
v: share, divide, split, distribute, allocate, apportion, give out
v: collect, gather, accumulate, assemble, amass, pile up, compile, stockpile
v: spread, scatter, disperse, distribute, circulate, broadcast, extend, diffuse
v: shake, tremble, quiver, shiver, quake, wobble, vibrate, shudder
v: turn, rotate, spin, revolve, twist, swivel, pivot, whirl
v: bend, curve, twist, arch, bow, flex, fold, crook
v: shine, glow, gleam, glitter, sparkle, glint, flash, beam, glisten
v: burn, blaze, flame, smoulder, scorch, char, singe, kindle
v: smell, sniff, scent, nose, detect, whiff
v: taste, sample, try, savour, sip, nibble
v: touch, feel, handle, stroke, pat, brush, finger, contact
v: hear, listen, catch, overhear, make out, attend
v: listen, hear, attend, pay attention, heed, hark
v: watch, observe, view, look at, monitor, survey, guard
v: sit, perch, settle, be seated, take a seat, rest
v: stand, rise, get up, stand up, be upright
v: lie, recline, lounge, sprawl, rest, stretch out
v: lie, fib, deceive, mislead, perjure oneself, bluff
v: pretend, feign, fake, act, bluff, put on, simulate
v: cheat, deceive, swindle, trick, con, defraud, dupe, double-cross
v: steal, rob, take, pinch, nick, swipe, pilfer, lift, embezzle
n: house, home, dwelling, residence, abode, property, building, lodging
n: home, house, residence, household, family, base, birthplace
n: building, structure, construction, edifice, premises, block, erection
n: room, chamber, space, area, compartment, hall
n: office, workplace, bureau, department, agency, headquarters, study
n: shop, store, outlet, boutique, retailer, market, supermarket, emporium
n: road, street, avenue, lane, way, route, highway, thoroughfare, drive
n: path, track, trail, footpath, walkway, lane, route, way
n: way, method, means, manner, approach, technique, mode, procedure, system
n: way, route, path, road, direction, course, passage
n: town, city, village, settlement, borough, municipality, township
n: country, nation, state, land, kingdom, realm, territory, republic
n: country, countryside, rural area, provinces, backwoods
n: world, earth, globe, planet, universe, creation
n: place, location, site, spot, position, point, area, venue, setting
n: area, region, zone, district, sector, territory, neighbourhood, locality
n: part, piece, portion, section, segment, fraction, bit, component, element
n: piece, bit, fragment, scrap, morsel, chunk, lump, slice, shred
n: whole, total, sum, entirety, aggregate, all, totality
n: group, set, collection, bunch, cluster, batch, crowd, band, party, gang
n: crowd, throng, mass, mob, horde, multitude, host, flock
n: team, side, squad, crew, group, unit, line-up
n: company, firm, business, enterprise, corporation, organisation, concern, establishment
n: company, companionship, fellowship, society, presence, friendship
n: business, trade, commerce, industry, dealings, transactions
n: job, work, occupation, profession, career, employment, post, position, role
n: job, task, chore, duty, assignment, errand, undertaking, project
n: work, labour, toil, effort, exertion, graft, drudgery
n: worker, employee, labourer, hand, member of staff, operative, workman
n: boss, manager, chief, head, director, employer, supervisor, leader
n: leader, head, chief, ruler, commander, captain, principal, director
n: friend, companion, mate, pal, comrade, ally, associate, chum
n: enemy, foe, opponent, adversary, rival, antagonist
n: person, individual, human, being, man, woman, soul, character
n: people, persons, individuals, folk, humans, public, population, citizens
n: child, kid, youngster, infant, toddler, baby, juvenile, minor
n: man, gentleman, male, fellow, guy, chap, bloke
n: woman, lady, female, girl, madam
n: family, relatives, relations, kin, household, clan, folks
n: parent, mother, father, guardian, carer
n: husband, partner, spouse, mate, consort
n: wife, partner, spouse, mate, consort
n: teacher, tutor, instructor, educator, lecturer, trainer, coach, master
n: student, pupil, learner, scholar, schoolchild, undergraduate, trainee
n: doctor, physician, medic, GP, consultant, surgeon, practitioner
n: customer, client, buyer, purchaser, patron, consumer, shopper
n: owner, proprietor, holder, possessor, landlord
n: expert, specialist, authority, professional, master, guru, connoisseur
n: beginner, novice, learner, newcomer, trainee, apprentice, recruit
n: idea, thought, notion, concept, plan, suggestion, theory, impression
n: plan, scheme, idea, proposal, project, design, strategy, programme
n: aim, goal, purpose, objective, target, intention, ambition, end
n: reason, cause, grounds, motive, basis, rationale, explanation, justification
n: result, outcome, consequence, effect, upshot, conclusion, product
n: effect, result, consequence, impact, influence, outcome
n: problem, difficulty, trouble, issue, snag, complication, obstacle, hitch
n: trouble, difficulty, problem, bother, worry, inconvenience, nuisance
n: answer, reply, response, solution, retort, reaction
n: question, query, enquiry, inquiry, problem, issue, matter
n: matter, issue, subject, topic, question, affair, point, concern
n: subject, topic, theme, matter, issue, question, point
n: fact, truth, reality, certainty, actuality, detail, point
n: detail, particular, point, item, element, feature, aspect, fact
n: information, data, facts, details, knowledge, intelligence, figures
n: news, information, report, word, story, announcement, bulletin
n: story, tale, account, narrative, report, anecdote, yarn, legend
n: book, volume, publication, title, work, tome, novel
n: paper, newspaper, journal, publication, daily, tabloid
n: letter, note, message, communication, missive, line
n: message, note, communication, word, memo, letter, notice, signal
n: word, term, expression, name, phrase, utterance
n: name, title, label, designation, term, tag, epithet
n: list, inventory, register, record, catalogue, schedule, roster, index
n: record, account, file, register, log, report, document, archive
n: document, paper, file, record, report, certificate, form
n: report, account, statement, review, record, summary, description
n: picture, image, painting, drawing, photograph, illustration, portrait, likeness
n: photograph, photo, picture, snapshot, image, shot, print
n: sign, signal, indication, mark, symbol, token, gesture, clue
n: mark, stain, spot, blemish, smudge, smear, blot
n: mistake, error, fault, slip, blunder, oversight, inaccuracy, gaffe
n: fault, defect, flaw, imperfection, weakness, failing, shortcoming
n: chance, opportunity, opening, possibility, occasion, prospect, break
n: chance, luck, fortune, fate, accident, coincidence, providence
n: risk, danger, hazard, peril, threat, gamble, jeopardy
n: danger, peril, hazard, risk, threat, menace, jeopardy
n: safety, security, protection, shelter, safe keeping
n: help, aid, assistance, support, backing, guidance, service, relief
n: support, backing, help, assistance, encouragement, aid, endorsement
n: advice, guidance, counsel, recommendation, suggestion, tip, pointer
n: rule, regulation, law, principle, guideline, standard, statute, order
n: law, rule, act, statute, regulation, decree, ordinance, legislation
n: right, entitlement, privilege, freedom, liberty, prerogative, claim
n: freedom, liberty, independence, autonomy, release, emancipation
n: power, strength, force, might, energy, potency, vigour
n: power, control, authority, command, rule, influence, dominance, sway
n: strength, power, might, force, muscle, toughness, stamina, vigour
n: energy, vigour, vitality, life, drive, spirit, zest, verve, stamina
n: effort, attempt, try, endeavour, exertion, struggle, labour, bid
n: success, achievement, triumph, victory, accomplishment, win, hit
n: failure, defeat, collapse, flop, disaster, fiasco, breakdown, washout
n: victory, win, triumph, success, conquest, mastery
n: defeat, loss, beating, rout, failure, setback, reverse
n: war, warfare, conflict, combat, fighting, hostilities, battle, struggle
n: fight, battle, struggle, clash, conflict, brawl, scuffle, skirmish
n: argument, dispute, quarrel, disagreement, row, debate, squabble, altercation
n: peace, calm, quiet, tranquillity, serenity, harmony, stillness, rest
n: noise, sound, din, racket, clamour, uproar, commotion, row, tumult
n: sound, noise, tone, note, ring, echo, resonance
n: silence, quiet, hush, stillness, calm, peace, tranquillity
n: light, brightness, glow, gleam, illumination, radiance, shine, brilliance
n: darkness, dark, gloom, shadow, night, blackness, murk
n: colour, hue, shade, tint, tone, tinge, pigment, dye
n: shape, form, figure, outline, contour, silhouette, profile, pattern
n: size, dimensions, magnitude, extent, scale, proportions, bulk, volume
n: amount, quantity, sum, total, volume, number, measure, mass
n: number, figure, digit, numeral, integer, quantity, amount, total
n: price, cost, charge, fee, rate, fare, value, expense, tariff
n: money, cash, funds, currency, capital, finance, wealth, riches
n: wealth, riches, fortune, money, prosperity, affluence, assets, means
n: gift, present, donation, offering, contribution, grant, bonus, handout
n: prize, award, reward, trophy, medal, honour, winnings
n: reward, prize, payment, bonus, recompense, return, payoff
n: pay, salary, wages, earnings, income, fee, payment, remuneration
n: cost, price, expense, charge, outlay, payment, expenditure
n: time, period, while, spell, stretch, span, term, interval, season
n: moment, instant, second, minute, flash, twinkling, jiffy
n: age, era, period, epoch, time, generation, days
n: end, finish, close, conclusion, completion, termination, finale, ending
n: start, beginning, outset, onset, opening, commencement, launch, inception
n: middle, centre, heart, core, midst, midpoint, hub
n: edge, border, rim, brink, margin, boundary, fringe, verge, side
n: top, summit, peak, crown, crest, tip, apex, head, pinnacle
n: bottom, base, foot, foundation, floor, bed, underside
n: side, edge, flank, face, surface, facet, aspect
n: front, face, facade, fore, head, lead, exterior
n: back, rear, reverse, end, tail, hind part
n: inside, interior, centre, core, middle, heart
n: outside, exterior, surface, face, shell, facade
n: line, row, queue, column, file, string, chain, series, rank
n: series, sequence, succession, chain, string, run, cycle, set
n: order, sequence, arrangement, series, succession, organisation, system
n: order, command, instruction, direction, demand, decree, directive
n: order, request, booking, reservation, commission, requisition
n: kind, sort, type, variety, class, category, form, species, genre
n: example, instance, case, illustration, sample, specimen, model
n: model, example, pattern, standard, ideal, template, prototype, paradigm
n: copy, duplicate, replica, reproduction, imitation, facsimile, clone
n: change, alteration, modification, adjustment, shift, transformation, variation, revision
n: increase, rise, growth, gain, expansion, boost, addition, surge
n: decrease, fall, drop, reduction, decline, cut, loss, dip
n: improvement, advance, progress, development, enhancement, upgrade, betterment
n: progress, advance, development, headway, growth, improvement, movement
n: move, movement, motion, action, step, gesture, shift
n: action, act, deed, step, move, measure, activity, operation
n: event, occasion, occurrence, happening, incident, affair, function
n: party, celebration, gathering, function, festivity, reception, do
n: meeting, gathering, assembly, conference, session, convention, get-together
n: journey, trip, voyage, tour, expedition, excursion, outing, travel
n: holiday, vacation, break, leave, time off, rest, recess
n: game, match, contest, competition, sport, tournament, play
n: fun, enjoyment, amusement, entertainment, pleasure, recreation, play
n: joy, happiness, delight, pleasure, gladness, elation, bliss, glee
n: sadness, sorrow, unhappiness, grief, misery, melancholy, gloom, despair
n: anger, rage, fury, wrath, annoyance, irritation, indignation, temper
n: fear, fright, terror, dread, alarm, panic, horror, anxiety
n: worry, anxiety, concern, unease, stress, nervousness, apprehension, trouble
n: love, affection, fondness, devotion, adoration, passion, tenderness
n: hate, hatred, loathing, dislike, hostility, animosity, contempt, aversion
n: hope, wish, desire, dream, aspiration, ambition, expectation, longing
n: wish, desire, hope, want, longing, yearning, craving, urge
n: pain, hurt, ache, soreness, discomfort, agony, suffering, torment
n: illness, sickness, disease, ailment, complaint, condition, malady, disorder
n: health, fitness, wellbeing, strength, vigour, condition, shape
n: body, figure, frame, physique, form, build, torso
n: mind, brain, intellect, head, intelligence, wits, reason, thought
n: feeling, emotion, sensation, sense, sentiment, impression, mood
n: mood, temper, humour, frame of mind, spirits, disposition, state
n: character, personality, nature, temperament, disposition, make-up, identity
n: skill, ability, talent, expertise, competence, aptitude, proficiency, knack
n: ability, capability, capacity, skill, talent, competence, power, faculty
n: knowledge, understanding, learning, wisdom, awareness, education, expertise
n: experience, knowledge, practice, skill, familiarity, background
n: experience, event, incident, episode, adventure, encounter, ordeal
n: education, schooling, teaching, training, instruction, learning, tuition
n: lesson, class, lecture, tutorial, session, period, seminar
n: test, examination, exam, assessment, trial, check, quiz, evaluation
n: study, research, investigation, analysis, examination, survey, review
n: research, study, investigation, enquiry, exploration, analysis, experimentation
n: science, knowledge, study, discipline, field
n: art, artwork, painting, craft, skill, creativity, design
n: music, song, tune, melody, harmony, composition
n: song, tune, melody, air, number, ballad, anthem, hymn
n: food, meal, nourishment, provisions, fare, cuisine, diet, grub
n: meal, dinner, lunch, breakfast, supper, feast, snack, repast
n: drink, beverage, refreshment, liquid
n: clothes, clothing, garments, dress, attire, outfit, wear, apparel
n: car, vehicle, motor, automobile, motorcar
n: machine, device, appliance, apparatus, engine, mechanism, gadget
n: tool, instrument, implement, device, utensil, gadget, appliance
n: thing, object, item, article, entity, device, gadget
n: things, belongings, possessions, stuff, goods, effects, gear, property
n: goods, products, merchandise, wares, stock, commodities
n: product, goods, item, commodity, merchandise, article, creation
n: material, substance, matter, stuff, fabric, cloth, medium
n: water, liquid, fluid, rain, sea
n: weather, climate, conditions, elements, forecast
n: rain, rainfall, shower, drizzle, downpour, deluge, precipitation
n: wind, breeze, gale, gust, draught, blast, air current
n: sun, sunshine, sunlight, daylight, warmth
n: sea, ocean, the deep, waters, main
n: hill, mount, rise, slope, elevation, knoll, hillock, mound
n: mountain, peak, summit, mount, alp, range, height
n: river, stream, brook, waterway, creek, tributary, channel
n: wood, forest, woodland, grove, copse, thicket, plantation
n: field, meadow, pasture, paddock, grassland, green, lea
n: garden, yard, grounds, lawn, park, plot, allotment
n: animal, creature, beast, being, brute
n: dog, hound, puppy, pup, canine, mutt
n: cat, kitten, feline, puss
n: noise, interference, static, disturbance
n: view, opinion, belief, idea, thought, judgement, attitude, position, conviction
n: opinion, view, belief, judgement, thought, feeling, verdict, assessment
n: belief, faith, conviction, trust, confidence, credo, principle
n: faith, belief, trust, confidence, conviction, religion, creed
n: trust, faith, confidence, belief, reliance, credence
n: truth, fact, reality, honesty, accuracy, veracity, sincerity
n: lie, untruth, falsehood, fib, fabrication, invention, deception
n: secret, mystery, confidence, enigma, puzzle, riddle
n: mystery, puzzle, enigma, riddle, secret, conundrum
n: choice, option, alternative, selection, pick, preference, decision
n: decision, choice, conclusion, verdict, judgement, ruling, resolution
n: agreement, deal, contract, arrangement, understanding, pact, treaty, bargain
n: promise, pledge, vow, word, guarantee, assurance, undertaking, oath
n: offer, proposal, bid, tender, proposition, suggestion
n: request, appeal, plea, application, demand, petition, entreaty
n: demand, request, call, claim, requirement, need, order
n: need, requirement, necessity, want, demand, essential, requisite
n: duty, responsibility, obligation, task, job, role, charge, function
n: purpose, aim, intention, goal, object, objective, function, point
n: use, purpose, function, point, application, value, utility, service
n: value, worth, merit, importance, usefulness, benefit, significance, price
n: benefit, advantage, gain, profit, good, help, plus, asset
n: advantage, benefit, asset, plus, gain, edge, upper hand, strength
n: disadvantage, drawback, downside, handicap, snag, weakness, minus, catch
n: loss, damage, harm, cost, deprivation, forfeit, waste
n: damage, harm, injury, destruction, loss, ruin, devastation, impairment
n: attack, assault, offensive, raid, strike, onslaught, invasion, charge
n: protection, defence, shelter, safeguard, cover, security, guard, shield
n: shelter, refuge, cover, protection, sanctuary, haven, retreat
n: system, method, structure, organisation, arrangement, network, scheme, set-up
n: method, way, means, technique, approach, procedure, process, system
n: process, procedure, method, operation, system, course, practice
n: style, manner, way, fashion, technique, approach, mode, method
n: fashion, style, trend, vogue, craze, fad, mode, look
n: habit, custom, practice, routine, tradition, convention, way, wont
n: tradition, custom, practice, convention, ritual, habit, institution
n: culture, civilisation, society, way of life, traditions, customs
n: society, community, the public, people, civilisation, nation
n: community, neighbourhood, society, population, people, residents, district
n: government, administration, authorities, state, regime, executive, cabinet
n: country, state, nation, power, land
n: army, military, troops, soldiers, forces, armed forces, militia
n: police, constabulary, force, officers, the law
n: crime, offence, wrongdoing, misdeed, felony, violation, transgression
n: criminal, offender, lawbreaker, villain, crook, felon, culprit, wrongdoer
n: punishment, penalty, sentence, discipline, retribution, fine
n: accident, crash, collision, mishap, disaster, misfortune, calamity
n: disaster, catastrophe, calamity, tragedy, misfortune, debacle, cataclysm
n: emergency, crisis, predicament, plight, danger, difficulty, exigency
n: rest, break, pause, respite, breather, lull, interval, relaxation
n: sleep, rest, nap, doze, slumber, snooze, siesta, kip
n: dream, vision, fantasy, daydream, nightmare, reverie, ambition
n: memory, recollection, remembrance, reminder, souvenir, reminiscence
n: future, outlook, prospects, expectations, tomorrow, time to come
n: past, history, background, yesterday, old days, antiquity
n: history, past, record, chronicle, annals, background, story
n: beauty, loveliness, attractiveness, charm, grace, elegance, splendour
n: quality, standard, grade, calibre, class, level, condition, merit
n: quality, feature, characteristic, attribute, property, trait, aspect
n: feature, characteristic, quality, attribute, aspect, trait, mark, point
n: level, standard, grade, degree, stage, rank, position, tier
n: stage, phase, step, point, period, level, juncture
n: step, stage, phase, move, measure, action, pace, stride
n: position, place, location, situation, spot, site, point
n: position, job, post, role, office, situation, appointment, placement
n: situation, circumstances, position, condition, state of affairs, case, predicament
n: condition, state, shape, form, order, fitness, repair
n: state, condition, shape, situation, mode, circumstances
n: space, room, area, capacity, expanse, extent, scope, gap
n: gap, space, opening, hole, break, crack, interval, breach
n: hole, opening, gap, cavity, hollow, pit, crater, puncture
n: bag, sack, pack, case, holdall, pouch, satchel, handbag
n: box, case, chest, container, crate, carton, package, packet
n: container, vessel, receptacle, holder, box, case, repository
n: present, gift, offering, donation
n: surprise, shock, astonishment, amazement, wonder, revelation, bombshell
n: interest, attention, curiosity, concern, notice, regard, involvement
n: interest, hobby, pastime, pursuit, activity, passion, leisure activity
n: attention, notice, heed, concentration, consideration, awareness, regard
n: care, attention, caution, prudence, thought, concern, vigilance
n: care, protection, charge, custody, keeping, supervision, safe keeping
r: quickly, fast, rapidly, swiftly, speedily, briskly, hastily, promptly
r: slowly, gradually, leisurely, unhurriedly, sluggishly, steadily
r: carefully, cautiously, attentively, meticulously, thoroughly, gingerly, warily
r: often, frequently, regularly, repeatedly, commonly, generally, usually
r: sometimes, occasionally, now and then, at times, from time to time, once in a while
r: always, constantly, continually, forever, invariably, perpetually, without exception
r: never, at no time, not ever, not once, nevermore
r: usually, normally, generally, typically, ordinarily, as a rule, mostly, commonly
r: very, extremely, really, highly, exceedingly, especially, exceptionally, truly, terribly
r: quite, fairly, rather, somewhat, reasonably, moderately, relatively, pretty
r: completely, totally, entirely, fully, wholly, utterly, absolutely, altogether
r: almost, nearly, practically, virtually, about, approximately, roughly, just about
r: really, truly, actually, genuinely, indeed, in fact, certainly
r: certainly, definitely, surely, undoubtedly, absolutely, positively, unquestionably
r: probably, likely, presumably, doubtless, in all likelihood, most likely
r: perhaps, maybe, possibly, conceivably, perchance
r: suddenly, abruptly, unexpectedly, all of a sudden, without warning, swiftly
r: soon, shortly, presently, before long, in a moment, any minute, quickly
r: immediately, at once, instantly, straight away, right away, forthwith, now
r: finally, eventually, at last, ultimately, lastly, in the end
r: again, once more, anew, afresh, another time
r: easily, effortlessly, readily, simply, smoothly, comfortably, with ease
r: well, satisfactorily, properly, correctly, effectively, competently, nicely
r: badly, poorly, inadequately, incompetently, terribly, awfully
r: happily, cheerfully, gladly, joyfully, merrily, contentedly, willingly
r: sadly, unhappily, sorrowfully, regretfully, unfortunately, dejectedly
r: quietly, silently, softly, noiselessly, calmly, peacefully, gently
r: loudly, noisily, deafeningly, boisterously, vociferously
r: clearly, plainly, obviously, evidently, distinctly, visibly, lucidly
r: especially, particularly, specially, notably, chiefly, mainly, above all
r: mainly, mostly, chiefly, largely, primarily, principally, predominantly, generally
r: also, too, as well, besides, in addition, moreover, furthermore, additionally
r: however, but, yet, nevertheless, nonetheless, still, though
r: therefore, so, thus, hence, consequently, accordingly, as a result
r: together, jointly, collectively, as one, in unison, side by side
r: alone, solo, single-handed, by oneself, independently, unaided
r: nearly, closely, almost, about, roughly, approximately
r: far, a long way, afar, distantly, widely, considerably, much
r: here, in this place, present, at hand, hither
r: everywhere, all over, throughout, far and wide, universally, all around
r: properly, correctly, appropriately, suitably, rightly, fittingly, aptly
r: hard, energetically, vigorously, diligently, industriously, strenuously, earnestly
r: gently, softly, lightly, tenderly, carefully, mildly, delicately
r: honestly, truthfully, sincerely, frankly, candidly, openly, genuinely
r: seriously, gravely, solemnly, earnestly, severely, badly
r: simply, merely, just, only, purely, plainly, solely
r: exactly, precisely, accurately, correctly, just, strictly, perfectly
r: generally, broadly, mostly, on the whole, overall, in general, as a rule
r: recently, lately, newly, freshly, of late, just now
r: currently, now, at present, presently, at the moment, nowadays, today
r: briefly, shortly, concisely, succinctly, momentarily, in short, in brief
v: provide, supply, give, furnish, offer, deliver, equip, contribute, grant
v: ensure, make sure, guarantee, secure, confirm, see to it, certify
v: achieve, accomplish, attain, reach, realise, gain, fulfil, complete, earn
v: deliver, provide, supply, hand over, bring, distribute, carry out, fulfil
v: implement, carry out, execute, apply, put into effect, enforce, enact, perform
v: maintain, keep up, preserve, sustain, continue, uphold, retain, conserve
v: maintain, claim, assert, insist, declare, contend, state, hold
v: support, back, help, assist, endorse, champion, uphold, sustain
v: require, need, demand, call for, necessitate, want, entail, involve
v: obtain, get, acquire, secure, gain, procure, achieve, attain
v: assess, evaluate, judge, appraise, gauge, estimate, rate, weigh up, review
v: evaluate, assess, appraise, judge, analyse, rate, measure, examine
v: analyse, examine, study, investigate, evaluate, inspect, scrutinise, interpret
v: review, examine, assess, reassess, survey, study, check, revise, look over
v: consider, think about, contemplate, weigh, ponder, examine, study, reflect on
v: discuss, talk about, debate, consider, examine, go over, confer, deliberate
v: propose, suggest, put forward, recommend, advance, submit, move, offer
v: confirm, verify, establish, prove, corroborate, validate, ratify, endorse
v: notify, inform, tell, advise, alert, warn, announce, report
v: submit, present, hand in, put forward, send, offer, tender, file
v: complete, finish, conclude, end, finalise, accomplish, wrap up, fulfil
v: update, revise, modernise, renew, refresh, upgrade, amend, bring up to date
v: issue, publish, release, distribute, circulate, put out, announce, send out
v: approve, endorse, sanction, authorise, accept, ratify, pass, agree to
v: reject, refuse, decline, dismiss, turn down, rebuff, veto, discard
v: launch, start, begin, introduce, initiate, open, unveil, set up
v: establish, set up, found, create, form, institute, start, build
v: organise, arrange, plan, coordinate, set up, manage, run, structure
v: coordinate, organise, arrange, harmonise, integrate, manage, synchronise
v: monitor, track, observe, watch, check, supervise, follow, survey
v: identify, recognise, pinpoint, determine, establish, spot, detect, name
v: determine, decide, establish, settle, fix, resolve, ascertain, find out
v: indicate, show, suggest, point to, signal, imply, reveal, denote
v: involve, include, entail, require, mean, comprise, contain, concern
v: concern, affect, involve, relate to, apply to, touch, regard
v: relate, connect, associate, link, correlate, tie in, refer
v: refer, mention, cite, allude, point to, relate, consult
v: present, show, display, introduce, give, offer, deliver, exhibit
v: represent, stand for, symbolise, depict, portray, embody, act for
v: consist, comprise, be made up of, contain, include, involve
v: contain, hold, include, comprise, incorporate, carry, accommodate
v: reduce, cut, lower, decrease, lessen, diminish, trim, curtail, minimise
v: expand, extend, grow, enlarge, broaden, increase, widen, develop, spread
v: extend, lengthen, prolong, expand, stretch, widen, enlarge, continue
v: limit, restrict, confine, curb, control, cap, constrain, restrain
v: focus, concentrate, centre, fix, direct, zero in on
v: emphasise, stress, highlight, underline, accentuate, insist on, spotlight
v: highlight, emphasise, stress, spotlight, accentuate, point up, feature
v: improve, enhance, boost, strengthen, upgrade, refine, develop, raise
v: strengthen, reinforce, fortify, bolster, toughen, consolidate, intensify, boost
v: weaken, undermine, impair, diminish, sap, erode, lessen, enfeeble
v: solve, resolve, settle, fix, sort out, work out, answer, clear up
v: avoid, evade, escape, dodge, shun, sidestep, steer clear of, bypass
v: reach, arrive at, get to, attain, achieve, come to, make
v: contact, get in touch with, reach, call, approach, communicate with, write to
v: respond, reply, answer, react, acknowledge, return
v: apply, use, employ, utilise, put into practice, exercise, implement
v: apply, request, put in, petition, appeal, bid, enquire
v: attend, go to, be present, visit, appear at, take part in
v: participate, take part, join in, engage, be involved, contribute, partake
v: contribute, give, donate, provide, supply, add, chip in
v: allocate, assign, allot, distribute, apportion, earmark, designate
v: assign, allocate, give, appoint, delegate, designate, allot, entrust
v: appoint, name, nominate, select, choose, assign, designate, install
v: employ, hire, engage, take on, recruit, appoint, retain
v: dismiss, sack, fire, discharge, lay off, let go
v: retire, withdraw, step down, resign, leave, quit
v: resign, quit, leave, step down, stand down, give up
v: negotiate, bargain, discuss, deal, arrange, settle, broker, mediate
v: invest, put money into, fund, finance, back, sponsor, stake
v: earn, make, get, gain, receive, bring in, collect, net, gross
v: charge, bill, invoice, levy, ask, demand
v: estimate, calculate, assess, judge, gauge, guess, reckon, approximate
v: calculate, compute, work out, reckon, count, determine, figure, estimate
v: predict, forecast, foresee, anticipate, project, prophesy, expect
v: plan, schedule, timetable, programme, arrange, book, slot in
v: delay, postpone, defer, hold up, put off, suspend, stall, adjourn
v: cancel, call off, abandon, scrap, abort, drop, revoke, annul
v: hurry, rush, hasten, speed, dash, race, accelerate, expedite
v: attach, fasten, connect, fix, join, secure, enclose, append
v: print, publish, issue, produce, reproduce, run off
v: store, keep, save, hold, file, stock, archive, deposit
v: delete, remove, erase, cut, cancel, eliminate, strike out, wipe
v: edit, revise, correct, amend, rewrite, redraft, adapt, polish
v: correct, rectify, fix, amend, adjust, remedy, put right, repair
v: format, arrange, lay out, design, set out, structure, style
v: sort, arrange, order, organise, classify, group, categorise, rank
v: search, look for, seek, hunt, explore, scan, probe, comb
v: select, choose, pick, opt for, single out, decide on
v: open, launch, start, begin, initiate, inaugurate
v: share, distribute, circulate, pass on, divulge, communicate
v: link, connect, join, associate, relate, tie, couple, attach
v: protect, secure, safeguard, guard, defend, preserve, shield
v: grow, increase, rise, climb, soar, escalate, mount, swell
v: fall, decline, drop, decrease, sink, slump, dip, plummet, tumble
v: rise, increase, go up, climb, mount, grow, soar, escalate
v: recover, improve, rally, revive, bounce back, pick up, recuperate, get better
v: suffer, endure, undergo, experience, bear, go through, tolerate, sustain
v: risk, chance, gamble, venture, endanger, hazard, jeopardise
v: struggle, strive, battle, labour, toil, grapple, contend, fight
v: wonder, ponder, speculate, think, reflect, puzzle, question, muse
v: doubt, question, distrust, mistrust, suspect, query, be uncertain
v: guess, estimate, suppose, reckon, surmise, conjecture, speculate, assume
v: assume, suppose, presume, take for granted, believe, imagine, expect, infer
v: mean, signify, indicate, denote, imply, suggest, convey, express
v: mean, intend, aim, plan, propose, design, purpose
v: express, convey, communicate, state, voice, articulate, put into words, show
v: mention, refer to, note, cite, bring up, touch on, allude to, raise
v: note, notice, observe, see, mark, register, perceive, remark
v: claim, assert, declare, maintain, allege, state, contend, profess
v: deny, refute, reject, contradict, dispute, rebut, disclaim, gainsay
v: admit, confess, acknowledge, concede, grant, own up, accept, allow
v: reveal, disclose, expose, uncover, divulge, unveil, show, make known
v: discover, find out, learn, ascertain, realise, determine, detect, unearth
v: realise, understand, recognise, see, grasp, appreciate, perceive, become aware
v: recognise, identify, know, place, distinguish, acknowledge, accept, spot
v: welcome, greet, receive, embrace, accept, hail, salute
v: celebrate, commemorate, mark, observe, honour, rejoice, party
v: enjoy, like, relish, savour, appreciate, delight in, love, revel in
v: prefer, favour, choose, like better, select, opt for, lean towards
v: tend, be inclined, be likely, be apt, lean, incline
v: matter, count, be important, signify, carry weight, make a difference
v: last, continue, go on, endure, persist, remain, survive, keep on
v: survive, live, last, endure, remain, persist, pull through, outlive
v: fit, suit, match, go with, correspond, conform, agree, tally
v: suit, fit, become, flatter, befit, satisfy, please
v: match, equal, correspond, fit, suit, go with, tally, agree
v: differ, vary, diverge, contrast, deviate, disagree, conflict
v: vary, change, differ, fluctuate, alter, diversify, range
v: compete, contend, vie, rival, challenge, struggle, take part
v: beat, defeat, overcome, conquer, outdo, surpass, trounce, vanquish
v: lead, head, top, be in front, precede, command
v: cooperate, collaborate, work together, join forces, team up, unite, combine
v: depend, rely, count, lean, bank, hang, hinge
v: wear, dress in, have on, put on, sport, don
v: dress, clothe, attire, garb, robe, kit out
v: cook, prepare, make, fix, heat, bake, roast, fry, boil
v: drive, steer, pilot, operate, motor, propel, chauffeur
v: travel, journey, tour, go, voyage, roam, wander, commute
v: fly, soar, glide, hover, float, wing, flutter, take off
v: swim, bathe, paddle, float, dive, wade
v: sing, chant, croon, hum, warble, serenade, chorus
v: dance, jig, sway, twirl, spin, skip, prance, caper
v: draw, sketch, trace, depict, outline, portray, illustrate, design
v: paint, colour, decorate, coat, illustrate, portray, depict
v: hang, suspend, dangle, sway, droop, drape, swing
v: wrap, cover, envelop, enfold, swathe, package, bundle, enclose
v: tie, bind, fasten, knot, secure, lash, tether, rope
v: mix, blend, combine, stir, merge, mingle, fuse, whisk
v: pour, flow, stream, spill, gush, run, tip, decant
v: fill in, complete, fill out, answer, enter
v: shut down, close, switch off, turn off, power down
v: switch on, turn on, start, activate, power up, boot
n: meeting, appointment, session, conference, consultation, interview, briefing
n: deadline, time limit, due date, cut-off, target date, closing date
n: budget, allowance, allocation, funds, resources, means, estimate, quota
n: profit, gain, earnings, return, income, surplus, yield, proceeds
n: revenue, income, earnings, turnover, receipts, proceeds, takings, sales
n: strategy, plan, policy, approach, scheme, tactic, programme, design
n: target, goal, aim, objective, purpose, end, mark, quota
n: growth, increase, expansion, development, rise, progress, enlargement
n: analysis, examination, study, review, evaluation, assessment, investigation, breakdown
n: staff, employees, personnel, workforce, workers, team, crew, human resources
n: client, customer, patron, buyer, consumer, user, account
n: service, help, assistance, aid, work, duty, favour
n: service, ceremony, rite, ritual, observance
n: project, scheme, plan, venture, undertaking, enterprise, assignment, programme
n: programme, schedule, timetable, agenda, plan, scheme, calendar
n: schedule, timetable, plan, programme, agenda, calendar, diary, itinerary
n: agenda, programme, schedule, list, plan, timetable
n: proposal, plan, suggestion, offer, scheme, recommendation, motion, bid
n: policy, plan, strategy, approach, procedure, line, guideline, code
n: issue, problem, matter, question, topic, concern, subject, point
n: issue, edition, number, copy, instalment, version
n: summary, synopsis, outline, abstract, overview, digest, résumé, recap
n: conclusion, end, close, finish, termination, completion, finale
n: conclusion, decision, judgement, verdict, deduction, inference, opinion
n: introduction, preface, foreword, opening, prologue, preamble, start
n: chapter, section, part, division, episode, stage, period
n: section, part, division, segment, portion, piece, unit, department
n: department, division, section, branch, unit, office, sector, bureau
n: branch, division, office, department, outlet, wing, arm, section
n: market, marketplace, mart, exchange, bazaar, fair, trade
n: industry, business, trade, field, sector, manufacturing, commerce
n: sale, deal, transaction, purchase, bargain, auction, trade
n: purchase, acquisition, buy, investment, asset, gain
n: deal, agreement, arrangement, transaction, contract, bargain, understanding
n: contract, agreement, deal, commitment, arrangement, settlement, undertaking
n: account, report, description, statement, story, version, record
n: account, bill, invoice, statement, tab, charge, balance
n: payment, fee, charge, instalment, remittance, settlement, sum, premium
n: debt, liability, obligation, arrears, dues, bill, loan
n: loan, advance, credit, mortgage, overdraft, debt, borrowing
n: tax, duty, levy, charge, toll, tariff, excise
n: fund, reserve, pool, kitty, supply, stock, collection, foundation
n: resources, assets, funds, capital, means, materials, reserves, supplies
n: stock, supply, store, reserve, stockpile, inventory, hoard, cache
n: supply, provision, stock, store, reserve, quantity, amount, delivery
n: delivery, consignment, shipment, dispatch, distribution, transportation
n: customer, consumer, user, client, shopper, buyer, patron
n: user, consumer, customer, operator, client
n: member, associate, fellow, participant, partner, subscriber
n: partner, associate, colleague, ally, collaborator, co-worker, companion
n: colleague, co-worker, associate, partner, collaborator, workmate
n: board, committee, council, panel, directorate, management, executive
n: committee, board, panel, council, commission, working party, group
n: authority, power, control, command, jurisdiction, influence, right
n: authority, expert, specialist, master, pundit, connoisseur, guru
n: responsibility, duty, obligation, liability, task, burden, role, charge
n: opportunity, chance, opening, possibility, occasion, break, time
n: challenge, test, trial, difficulty, problem, task, obstacle, hurdle
n: obstacle, barrier, hurdle, difficulty, obstruction, impediment, hindrance, snag
n: solution, answer, resolution, remedy, key, explanation, fix, way out
n: improvement, enhancement, upgrade, advance, gain, development, refinement
n: success, triumph, achievement, accomplishment, victory, prosperity, hit, sensation
n: performance, achievement, accomplishment, execution, conduct, work, results
n: performance, show, production, presentation, act, recital, concert, staging
n: presentation, talk, speech, lecture, demonstration, display, show, address
n: speech, address, talk, lecture, presentation, oration, sermon
n: discussion, debate, conversation, talk, consultation, exchange, dialogue, deliberation
n: conversation, talk, chat, discussion, dialogue, exchange, conference
n: email, message, mail, note, e-mail
n: letter, character, symbol, sign, figure
n: page, sheet, leaf, folio, side
n: table, chart, grid, list, diagram, schedule, tabulation
n: chart, graph, diagram, table, plot, figure, map
n: file, folder, document, record, dossier, archive
n: copy, print, duplicate, version, edition, reproduction
n: version, edition, form, variant, adaptation, interpretation, rendering
n: draft, outline, sketch, plan, version, rough, first version
n: total, sum, whole, aggregate, amount, figure, tally, grand total
n: average, mean, norm, standard, median, middle, par
n: rate, speed, pace, tempo, velocity, frequency
n: rate, charge, price, fee, tariff, cost, figure
n: percentage, proportion, share, part, portion, fraction, ratio
n: share, portion, part, quota, allocation, ration, cut, percentage
n: figure, number, digit, statistic, amount, total, sum
n: figure, shape, form, body, outline, physique, silhouette
n: trend, tendency, movement, direction, pattern, drift, inclination, course
n: pattern, design, motif, arrangement, figure, device, decoration
n: pattern, system, arrangement, order, method, sequence, plan
n: aspect, feature, side, facet, element, angle, point, dimension
n: factor, element, component, aspect, feature, consideration, influence, cause
n: element, component, part, ingredient, factor, constituent, unit, piece
n: item, article, object, thing, piece, unit, entry, point
n: unit, component, part, element, section, module, segment, item
n: equipment, apparatus, gear, kit, tools, machinery, materials, supplies
n: software, program, application, app, code, package
n: computer, PC, machine, laptop, desktop, processor
n: network, system, web, grid, structure, net, organisation
n: link, connection, relationship, association, tie, bond, relation
n: relationship, connection, association, link, relation, bond, tie, rapport
n: contact, touch, communication, connection, correspondence
n: contact, connection, acquaintance, associate, source
n: access, entry, admission, admittance, entrance, approach, way in
n: entrance, entry, way in, door, doorway, gate, access, opening
n: exit, way out, door, outlet, egress, departure
n: border, boundary, frontier, edge, limit, margin, perimeter, line
n: limit, boundary, maximum, ceiling, restriction, cap, threshold, extent
n: range, scope, extent, reach, span, spread, variety, selection
n: variety, range, assortment, selection, mixture, diversity, choice, array
n: mixture, blend, combination, mix, compound, medley, assortment, amalgam
n: balance, equilibrium, stability, harmony, proportion, symmetry, evenness
n: comfort, ease, cosiness, relaxation, wellbeing, contentment, luxury
n: pleasure, enjoyment, delight, satisfaction, gratification, happiness, joy, fun
n: satisfaction, contentment, pleasure, fulfilment, gratification, happiness, pride
n: pride, satisfaction, self-respect, dignity, honour, self-esteem
n: shame, embarrassment, disgrace, humiliation, guilt, remorse, dishonour
n: respect, regard, esteem, admiration, honour, reverence, deference
n: honour, respect, esteem, distinction, credit, dignity, integrity, glory
n: reputation, name, standing, status, character, renown, prestige, image
n: status, standing, rank, position, prestige, importance, level, grade
n: courage, bravery, valour, boldness, daring, nerve, pluck, heroism, guts
n: confidence, assurance, self-belief, poise, certainty, trust, faith, aplomb
n: patience, tolerance, endurance, restraint, forbearance, composure, calm
n: kindness, goodness, generosity, compassion, consideration, warmth, charity
n: generosity, kindness, charity, liberality, bounty, benevolence, magnanimity
n: wisdom, knowledge, understanding, insight, judgement, sense, intelligence, prudence
n: intelligence, cleverness, brains, intellect, wit, ability, acumen, understanding
n: humour, comedy, wit, fun, amusement, jokes, levity, drollery
n: joke, jest, quip, gag, pun, wisecrack, prank, witticism
n: trick, prank, ruse, deception, ploy, stunt, joke, hoax
n: skill, technique, art, craft, expertise, mastery, flair, proficiency
n: talent, gift, flair, aptitude, ability, genius, knack, skill
n: genius, brilliance, intellect, talent, flair, gift, ability
n: effort, work, exertion, labour, energy, strain, struggle, application
n: hurry, rush, haste, speed, urgency, hustle, bustle
n: speed, pace, rate, velocity, swiftness, rapidity, quickness, tempo
n: delay, hold-up, wait, postponement, setback, interruption, lag, deferment
n: end, aim, goal, purpose, object, intention, objective
a: significant, important, notable, noteworthy, considerable, meaningful, momentous, major, substantial
a: considerable, substantial, significant, sizeable, appreciable, large, great, marked
a: numerous, many, several, various, countless, plentiful, abundant, multiple, a lot of
a: several, some, a few, various, sundry, assorted, a number of
a: relevant, pertinent, applicable, related, appropriate, apt, material, germane, to the point
a: appropriate, suitable, proper, fitting, apt, right, correct, apposite, seemly
a: suitable, appropriate, fit, fitting, proper, right, apt, acceptable, satisfactory
a: effective, successful, efficient, productive, powerful, useful, potent, effectual
a: efficient, effective, productive, competent, capable, economical, streamlined, organised
a: accurate, exact, precise, correct, true, right, faithful, meticulous
a: precise, exact, accurate, specific, particular, careful, meticulous, strict
a: detailed, thorough, comprehensive, full, exhaustive, minute, elaborate, in-depth
a: brief, short, concise, quick, succinct, compact, terse, momentary
a: concise, brief, succinct, short, terse, compact, pithy, to the point
a: positive, optimistic, hopeful, confident, upbeat, encouraging, constructive, favourable
a: negative, pessimistic, unfavourable, adverse, gloomy, harmful, defeatist, damaging
a: favourable, positive, good, encouraging, promising, beneficial, advantageous, approving
a: likely, probable, possible, expected, anticipated, plausible, liable, apt
a: unlikely, improbable, doubtful, implausible, remote, slim, far-fetched
a: obvious, apparent, evident, plain, clear, manifest, noticeable, visible
a: crucial, vital, critical, essential, key, central, decisive, pivotal, important
a: essential, vital, necessary, crucial, indispensable, key, fundamental, basic
a: basic, fundamental, essential, elementary, simple, rudimentary, primary, core
a: fundamental, basic, essential, primary, key, central, underlying, elementary
a: primary, main, chief, principal, prime, leading, key, first, initial
a: major, main, chief, leading, important, significant, principal, key, serious
a: minor, small, slight, unimportant, trivial, petty, secondary, lesser
a: serious, severe, grave, critical, acute, dangerous, alarming, major
a: severe, harsh, strict, stern, extreme, intense, serious, acute, drastic
a: strict, stern, severe, harsh, firm, rigid, rigorous, exacting
a: extreme, intense, severe, acute, utmost, excessive, drastic, radical
a: intense, extreme, strong, powerful, acute, deep, fierce, profound, keen
a: gentle, mild, soft, tender, kind, light, moderate, delicate
a: mild, gentle, moderate, soft, slight, temperate, calm, balmy
a: moderate, average, medium, reasonable, fair, modest, middling, restrained
a: reasonable, fair, sensible, rational, logical, sound, moderate, acceptable
a: logical, rational, reasonable, sound, coherent, consistent, valid, sensible
a: consistent, constant, steady, regular, stable, reliable, uniform, unchanging
a: stable, steady, secure, firm, fixed, solid, sound, balanced
a: flexible, adaptable, versatile, adjustable, elastic, supple, pliable, accommodating
a: rigid, stiff, inflexible, firm, hard, unbending, strict, set
a: active, busy, energetic, lively, dynamic, vigorous, involved, engaged
a: passive, inactive, inert, submissive, docile, unresponsive, idle
a: independent, self-reliant, self-sufficient, autonomous, free, separate, sovereign
a: dependent, reliant, conditional, subject, contingent, relying
a: original, novel, new, fresh, innovative, creative, inventive, unique, first
a: creative, imaginative, inventive, original, artistic, innovative, inspired, resourceful
a: unique, single, sole, distinctive, special, individual, unparalleled, one-off
a: typical, characteristic, representative, normal, standard, usual, classic, average
a: traditional, conventional, customary, established, classic, time-honoured, orthodox
a: current, present, existing, contemporary, ongoing, prevailing, latest, up-to-date
a: future, coming, forthcoming, prospective, impending, later, eventual, upcoming
a: former, previous, earlier, past, prior, ex, old, one-time
a: potential, possible, likely, prospective, budding, would-be, probable
a: actual, real, true, genuine, factual, concrete, existing, confirmed
a: physical, bodily, material, concrete, tangible, real, solid, corporeal
a: mental, intellectual, cerebral, psychological, cognitive, rational
a: emotional, passionate, sentimental, moving, touching, poignant, heartfelt, tearful
a: personal, private, individual, own, intimate, particular, subjective
a: social, communal, public, collective, community, civic, societal
a: economic, financial, commercial, monetary, fiscal, budgetary, profitable
a: political, governmental, civic, public, official, state, legislative
a: legal, lawful, legitimate, valid, permitted, authorised, constitutional, licit
a: illegal, unlawful, illicit, criminal, prohibited, forbidden, banned, outlawed
a: official, formal, authorised, approved, sanctioned, legitimate, licensed
a: formal, official, ceremonial, proper, conventional, correct, stiff
a: informal, casual, relaxed, unofficial, easy-going, friendly, colloquial
a: local, regional, neighbourhood, community, district, provincial, nearby
a: national, countrywide, nationwide, state, domestic, public, general
a: international, global, worldwide, universal, foreign, multinational, overseas
a: foreign, overseas, international, alien, exotic, external, unfamiliar
a: domestic, home, household, family, internal, national, inland
a: central, middle, main, key, core, inner, chief, principal
a: open, available, accessible, public, unrestricted, free
a: available, accessible, obtainable, at hand, on hand, ready, free, vacant
a: aware, conscious, informed, mindful, alert, knowledgeable, cognisant, alive to
a: unaware, ignorant, oblivious, unconscious, uninformed, unsuspecting
a: responsible, accountable, liable, answerable, in charge, reliable, sensible, mature
a: responsible, sensible, reliable, trustworthy, dependable, mature, conscientious
a: confident, sure, certain, self-assured, assured, positive, poised, bold
a: nervous, anxious, worried, tense, uneasy, apprehensive, jittery, on edge, edgy
a: worried, anxious, concerned, troubled, uneasy, nervous, fretful, disturbed
a: relaxed, calm, easy-going, at ease, laid-back, unworried, comfortable, carefree
a: comfortable, cosy, snug, relaxed, pleasant, homely, restful, easy
a: uncomfortable, awkward, uneasy, embarrassed, ill at ease, painful, cramped
a: embarrassed, awkward, ashamed, self-conscious, mortified, sheepish, red-faced
a: excited, thrilled, eager, enthusiastic, elated, animated, keen, delighted
a: bored, uninterested, weary, fed up, jaded, listless, restless
a: interested, curious, keen, attentive, engaged, absorbed, fascinated, intrigued
a: curious, inquisitive, interested, nosy, questioning, intrigued
a: surprised, astonished, amazed, startled, shocked, stunned, astounded, taken aback
a: disappointed, let down, dissatisfied, disheartened, dismayed, saddened, crestfallen
a: satisfied, content, pleased, happy, fulfilled, gratified, contented
a: jealous, envious, covetous, resentful, possessive, green-eyed, bitter
a: confused, puzzled, bewildered, baffled, perplexed, muddled, mystified, lost
a: confusing, puzzling, baffling, bewildering, perplexing, unclear, complicated
a: famous, celebrated, star, big-name
a: honest, fair, decent, upright, principled, ethical, moral, honourable
a: wicked, evil, bad, sinful, immoral, villainous, corrupt, malicious
a: good, virtuous, moral, righteous, upright, honourable, decent, worthy
a: gentle, kind, mild, tender, soft-hearted, caring, sweet, meek
a: tough, hardy, resilient, robust, sturdy, strong, rugged, durable
a: durable, lasting, hard-wearing, strong, tough, sturdy, long-lasting, robust
a: fragile, delicate, breakable, brittle, flimsy, frail, dainty
a: quiet, peaceful, tranquil, restful, calm, still, undisturbed, serene
a: noisy, loud, rowdy, boisterous, clamorous, rackety, deafening
a: crowded, packed, busy, full, jammed, congested, teeming, thronged
a: deserted, empty, abandoned, desolate, vacant, uninhabited, lonely, forsaken
a: remote, distant, far, isolated, faraway, outlying, secluded, inaccessible
a: near, close, nearby, adjacent, neighbouring, at hand, handy, local
a: distant, far, remote, faraway, far-off, outlying, removed
a: huge, enormous, vast, colossal, massive, gigantic, immense, tremendous, mammoth
a: tiny, minute, minuscule, microscopic, infinitesimal, wee, diminutive
a: great, large, big, huge, considerable, immense, vast, extensive
a: great, excellent, fantastic, terrific, marvellous, wonderful, brilliant, superb
a: great, important, eminent, distinguished, notable, famous, prominent, leading
a: brilliant, bright, dazzling, vivid, intense, shining, glowing, sparkling
a: brilliant, clever, gifted, talented, intelligent, able, exceptional, genius
a: brilliant, superb, excellent, wonderful, outstanding, magnificent, marvellous
a: sweet, sugary, sweetened, syrupy, honeyed, saccharine
a: sweet, kind, charming, lovely, pleasant, agreeable, gentle, endearing
a: bitter, sour, sharp, acid, tart, acrid, astringent
a: bitter, resentful, embittered, sour, hostile, rancorous, aggrieved
a: fresh, new, recent, crisp, raw, natural, unused, clean
a: stale, old, dry, mouldy, musty, off, hard, flat
a: rotten, decayed, decomposed, mouldy, putrid, rancid, bad, off
a: raw, uncooked, fresh, natural, crude, unprocessed, unrefined
a: famous, iconic, legendary, celebrated
a: angry, annoyed, irritated, exasperated, vexed, displeased, put out
a: delighted, thrilled, overjoyed, elated, ecstatic, pleased, glad, jubilant
a: hungry, starving, famished, ravenous, peckish, empty
a: thirsty, parched, dry, dehydrated, gasping
a: full, satisfied, replete, sated, stuffed, gorged
a: awake, conscious, alert, wakeful, wide awake, up, astir
a: asleep, sleeping, dozing, napping, slumbering, dormant, unconscious
a: alive, living, live, existing, breathing, surviving, animate
a: dead, deceased, lifeless, departed, late, extinct, gone
a: dead, flat, lifeless, inactive, inoperative, not working, defunct
r: carefully, closely, thoroughly, attentively, painstakingly, precisely, rigorously
r: significantly, considerably, substantially, greatly, markedly, notably, appreciably
r: slightly, a little, a bit, somewhat, marginally, moderately, faintly
r: greatly, very much, considerably, enormously, hugely, immensely, highly, vastly
r: entirely, completely, totally, wholly, fully, utterly, absolutely, altogether
r: largely, mostly, mainly, chiefly, primarily, principally, generally, predominantly
r: directly, straight, immediately, personally, firsthand, plainly
r: effectively, efficiently, successfully, productively, well, competently
r: currently, presently, now, at the moment, today, at present
r: previously, before, earlier, formerly, once, beforehand, hitherto
r: later, afterwards, subsequently, next, then, after, eventually, in time
r: early, ahead, beforehand, in advance, prematurely, betimes
r: late, behind, belatedly, tardily, overdue
r: often, many times, again and again, frequently, regularly, repeatedly
r: rarely, seldom, infrequently, hardly ever, occasionally, sporadically
r: fortunately, luckily, happily, thankfully, mercifully, providentially
r: unfortunately, sadly, unhappily, unluckily, regrettably, alas
r: obviously, clearly, plainly, evidently, apparently, patently, of course
r: apparently, seemingly, evidently, ostensibly, outwardly, supposedly
r: especially, notably, particularly, specially, exceptionally, extraordinarily
r: roughly, approximately, about, around, nearly, more or less, circa
r: instead, rather, alternatively, in its place, in lieu
r: otherwise, else, or, alternatively, differently
r: meanwhile, in the meantime, at the same time, simultaneously, for now
r: indeed, in fact, actually, really, truly, certainly, of course
r: nevertheless, nonetheless, however, still, yet, even so, all the same
r: furthermore, moreover, besides, also, in addition, what is more, additionally
v: argue, contend, maintain, claim, assert, reason, hold, insist
v: demonstrate, show, prove, establish, display, illustrate, exhibit, confirm
v: illustrate, show, demonstrate, depict, exemplify, explain, clarify, picture
v: define, explain, describe, specify, determine, establish, characterise, set out
v: emphasise, underline, stress, insist, accentuate
v: examine, inspect, study, look at, investigate, explore, analyse, scrutinise
v: investigate, examine, explore, research, study, look into, inquire into, probe
v: explore, investigate, examine, look into, research, survey, tour, travel
v: observe, notice, see, watch, witness, note, perceive, regard
v: occur, happen, take place, arise, come about, appear, crop up, transpire
v: arise, occur, emerge, appear, crop up, happen, develop, result
v: emerge, appear, come out, surface, arise, materialise, become known
v: appear, emerge, show, turn up, arrive, surface, materialise, seem
v: disappear, vanish, fade, evaporate, melt away, recede, go, depart
v: exist, be, live, survive, occur, remain, subsist, endure
v: generate, produce, create, cause, make, bring about, give rise to, yield
v: obtain, collect, gather, acquire, receive, earn
v: perform, carry out, do, execute, accomplish, complete, conduct, undertake
v: perform, act, play, present, stage, put on, appear, sing
v: conduct, carry out, perform, run, manage, direct, organise, hold
v: undertake, take on, assume, tackle, carry out, embark on, accept, attempt
v: handle, deal with, manage, cope with, tackle, treat, control, address
v: address, deal with, tackle, attend to, handle, take up, approach
v: treat, handle, deal with, regard, consider, use, behave towards
v: treat, cure, heal, nurse, tend, attend to, medicate, remedy
v: cure, heal, remedy, treat, restore, mend, fix, rectify
v: affect, impact, influence, alter, modify, sway, transform
v: influence, affect, sway, shape, guide, persuade, induce, steer
v: transform, change, convert, alter, revolutionise, reshape, remodel, metamorphose
v: convert, change, turn, transform, adapt, alter, switch
v: adapt, adjust, modify, alter, change, convert, tailor, fit
v: adjust, alter, modify, change, regulate, tune, tweak, fine-tune
v: reflect, mirror, echo, show, reveal, indicate, express, demonstrate
v: reflect, think, consider, ponder, contemplate, muse, deliberate
v: compare, contrast, set against, weigh, match, liken, equate
v: combine, merge, join, unite, blend, mix, amalgamate, integrate
v: divide, split, separate, part, share, cut, break up, partition
v: rank, grade, rate, classify, order, place, position, sort
v: measure, assess, rate, gauge, quantify, evaluate, judge, appraise
v: select, nominate, appoint, elect, vote for, name
v: vote, elect, choose, ballot, poll, opt
v: rule, govern, reign, control, lead, administer, command, run
v: serve, help, assist, work for, aid, attend to, wait on
v: serve, act, function, work, operate
v: function, work, operate, run, go, perform, act, serve
v: operate, run, work, function, use, control, manage, drive
v: fail, break down, stop working, malfunction, crash, give out
v: crash, collide, smash, hit, bump, run into, plough into
v: hurry up, get a move on, speed up, look sharp, buck up
v: arrive, come, reach, get to, turn up, land, appear, show up
v: leave, go, depart, exit, quit, abandon, desert, set off
v: abandon, leave, desert, forsake, quit, drop, give up, discontinue
v: quit, leave, give up, abandon, stop, resign, drop out
v: escape, flee, get away, run away, break free, bolt, abscond
v: chase, pursue, follow, hunt, track, run after, tail
v: hunt, search, seek, look for, pursue, chase, stalk, track
v: wish, want, desire, long, hope, crave, fancy, yearn
v: enjoy oneself, have fun, have a good time, make merry
a: able, capable, competent, skilled, proficient, accomplished, qualified, adept, talented
a: unable, incapable, powerless, unfit, unqualified, incompetent, impotent
a: skilful, skilled, expert, adept, deft, proficient, masterly, accomplished, dexterous
a: clumsy, awkward, bumbling, gauche, ungainly, inept, heavy-handed, maladroit
a: graceful, elegant, poised, fluid, smooth, agile, nimble, lithe
a: ambitious, aspiring, driven, determined, enterprising, motivated, go-getting, purposeful
a: determined, resolute, firm, purposeful, persistent, tenacious, single-minded, dogged, steadfast
a: stubborn, obstinate, headstrong, wilful, inflexible, pig-headed, intransigent, mulish
a: obedient, compliant, dutiful, submissive, docile, law-abiding, deferential, biddable
a: naughty, disobedient, mischievous, badly behaved, unruly, wayward, impish, bad
a: cheerful, bright, sunny, jolly, lively, merry, buoyant, chirpy, good-humoured
a: gloomy, depressed, downcast, morose, sullen, despondent, dismal, low-spirited
a: grumpy, bad-tempered, irritable, cantankerous, crabby, crotchety, tetchy, surly, sulky
a: thoughtful, considerate, attentive, kind, caring, solicitous, helpful, obliging
a: thoughtful, pensive, reflective, contemplative, meditative, absorbed, introspective
a: selfish, self-centred, egotistic, greedy, inconsiderate, thoughtless, mean
a: greedy, grasping, avaricious, acquisitive, materialistic, insatiable, gluttonous
a: honest, scrupulous, conscientious, principled, upstanding, incorruptible
a: trustworthy, reliable, dependable, honest, loyal, faithful, responsible, sound
a: suspicious, distrustful, wary, sceptical, doubtful, mistrustful, cynical, questioning
a: suspicious, questionable, dubious, shady, fishy, suspect, doubtful, irregular
a: famous, infamous, notorious, disreputable, scandalous
a: gentle, placid, docile, tame, mild, calm, easy-going
a: wild, untamed, feral, savage, ferocious, fierce, undomesticated
a: wild, uncontrolled, unruly, riotous, rowdy, frantic, disorderly, chaotic
a: fierce, ferocious, savage, vicious, violent, aggressive, brutal, wild
a: violent, brutal, savage, vicious, aggressive, rough, destructive, fierce
a: peaceful, non-violent, peaceable, calm, gentle, conciliatory, amicable
a: sensitive, delicate, tender, sore, raw, painful, touchy
a: sensitive, understanding, perceptive, sympathetic, aware, intuitive, tactful
a: tactful, diplomatic, discreet, sensitive, delicate, considerate, polite, careful
a: tactless, undiplomatic, insensitive, blunt, indiscreet, clumsy, thoughtless
a: rare, precious, unusual, valuable, special, uncommon
a: ordinary, unexceptional, commonplace, mundane, everyday, humdrum, prosaic
a: extraordinary, remarkable, exceptional, unusual, amazing, phenomenal, astonishing, unprecedented
a: incredible, unbelievable, amazing, astonishing, extraordinary, remarkable, implausible
a: impressive, striking, remarkable, imposing, grand, splendid, magnificent, awe-inspiring
a: magnificent, splendid, grand, impressive, superb, majestic, glorious, stately
a: elegant, stylish, graceful, chic, refined, smart, sophisticated, polished
a: shabby, tatty, scruffy, worn, threadbare, ragged, run-down, dilapidated
a: smart, neat, well-dressed, elegant, stylish, trim, dapper, spruce
a: casual, informal, relaxed, easy-going, laid-back, offhand, nonchalant
a: sudden, quick, hurried, rushed, hasty, swift, abrupt
a: hasty, hurried, rushed, quick, rapid, careless, rash, impetuous
a: rash, reckless, impulsive, hasty, impetuous, foolhardy, careless, unwise
a: famous, renowned, eminent, prestigious, esteemed, respected
a: respected, esteemed, admired, valued, revered, honoured, reputable, well thought of
a: respectable, decent, proper, reputable, honourable, upright, worthy, presentable
a: shameful, disgraceful, scandalous, dishonourable, outrageous, deplorable, contemptible
a: outrageous, shocking, scandalous, disgraceful, appalling, monstrous, atrocious, offensive
a: shocking, appalling, horrifying, disturbing, dreadful, outrageous, staggering, sickening
a: horrible, horrid, dreadful, awful, ghastly, hideous, nasty, frightful
a: frightening, scary, terrifying, alarming, chilling, eerie, menacing, hair-raising
a: spooky, eerie, creepy, uncanny, ghostly, weird, sinister, unearthly
a: sinister, menacing, threatening, ominous, evil, dark, forbidding
a: harmful, damaging, detrimental, injurious, hurtful, dangerous, destructive, bad
a: harmless, safe, innocuous, inoffensive, benign, gentle, mild
a: beneficial, helpful, advantageous, good, valuable, useful, favourable, wholesome
a: profitable, lucrative, money-making, rewarding, productive, worthwhile, gainful
a: worthwhile, valuable, useful, rewarding, beneficial, productive, constructive, profitable
a: pointless, futile, useless, vain, meaningless, senseless, worthless, unproductive
a: meaningful, significant, important, worthwhile, valuable, purposeful, relevant
a: proper, correct, right, appropriate, suitable, fitting, accepted, conventional
a: improper, unsuitable, inappropriate, wrong, unseemly, indecent, unfitting
a: equal, identical, equivalent, the same, even, level, uniform, matching
a: unequal, uneven, unbalanced, disproportionate, different, unfair, lopsided
a: total, complete, entire, whole, full, overall, comprehensive, utter
a: utter, complete, total, absolute, sheer, downright, thorough, outright
a: pure, clean, clear, fresh, unpolluted, untainted, unadulterated, natural
a: pure, sheer, utter, absolute, complete, total, simple
a: dirty, polluted, contaminated, impure, foul, infected, tainted
a: busy, industrious, hard-working, diligent, assiduous, conscientious, energetic
a: hard-working, diligent, industrious, conscientious, tireless, assiduous, dedicated
a: dedicated, committed, devoted, loyal, faithful, enthusiastic, wholehearted
a: keen, enthusiastic, eager, avid, ardent, fervent, passionate, zealous
a: passionate, intense, fervent, ardent, emotional, heated, fiery, impassioned
a: cool, cold, distant, aloof, reserved, unfriendly, frosty, standoffish
a: cool, calm, composed, collected, unflappable, self-possessed, relaxed, level-headed
a: warm, friendly, cordial, kindly, affectionate, hospitable, welcoming, genial
a: warm, mild, balmy, pleasant, temperate, summery, sunny
a: cosy, snug, comfortable, warm, homely, intimate, relaxed
a: plain, simple, ordinary, unadorned, basic, modest, homely, austere
a: fancy, elaborate, ornate, decorative, decorated, showy, ostentatious, intricate
a: grand, splendid, magnificent, impressive, imposing, stately, majestic, palatial
a: humble, modest, simple, lowly, ordinary, meek, unpretentious
a: vast, huge, enormous, immense, extensive, boundless, infinite, endless
a: limited, restricted, finite, small, narrow, confined, meagre, minimal
a: endless, infinite, boundless, unlimited, ceaseless, eternal, perpetual, interminable
a: temporary, short-term, brief, provisional, passing, fleeting, momentary, transient
a: permanent, lasting, enduring, everlasting, fixed, stable, perpetual, abiding
a: brief, momentary, fleeting, short-lived, passing, temporary, transitory
a: frequent, recurring, regular, habitual, constant, continual, periodic
a: daily, everyday, day-to-day, routine, quotidian, regular
a: annual, yearly, once a year, twelve-monthly
a: equal, fair, even-handed, balanced, impartial, just
a: whole, healthy, sound, fit, well, hale
a: obvious, blatant, flagrant, glaring, conspicuous, barefaced, overt
a: subtle, delicate, slight, faint, understated, muted, fine, nuanced
a: sudden, dramatic, striking, spectacular, sensational, marked
a: dramatic, theatrical, sensational, spectacular, striking, exciting, vivid, extreme
a: calm, unexcited, cool, level-headed, quiet, untroubled, sober
a: tense, strained, stressful, nerve-racking, fraught, uneasy, edgy, charged
a: stressful, demanding, difficult, tense, taxing, pressured, harrowing, worrying
a: relaxing, restful, calming, soothing, peaceful, tranquil, leisurely
a: noisy, deafening, ear-splitting, piercing, thunderous
a: hopeful, optimistic, confident, positive, expectant, buoyant, sanguine
a: hopeless, desperate, despairing, pessimistic, futile, impossible, forlorn, despondent
a: desperate, frantic, despairing, urgent, hopeless, dire, critical, drastic
a: anxious, eager, keen, impatient, longing
a: famous, best-selling, popular, hit, successful
a: innocent, harmless, inoffensive, blameless, pure
a: guilty, ashamed, remorseful, sorry, repentant, contrite, shamefaced
a: alert, watchful, vigilant, attentive, observant, wide awake, on the lookout, sharp
a: careless, absent-minded, forgetful, distracted, inattentive, vague, scatterbrained
a: forgetful, absent-minded, vague, scatty, distracted, oblivious
a: famous, popular, legendary, star-studded
a: fun, enjoyable, entertaining, amusing, pleasant, lively, diverting
a: playful, frisky, lively, sportive, mischievous, cheeky, frolicsome, fun-loving
a: lucky, happy, fortunate, opportune, timely, favourable, auspicious
a: grateful, glad, pleased, relieved, thankful
a: relieved, reassured, comforted, thankful, glad, eased, consoled
a: lonely, isolated, cut off, alone, abandoned, rejected, unloved
a: sociable, outgoing, gregarious, friendly, convivial, extrovert, companionable
a: quiet, reserved, introverted, withdrawn, shy, retiring, silent, taciturn
a: talkative, chatty, garrulous, communicative, voluble, loquacious, gossipy, wordy
a: famous, high-profile, prominent, public
v: love, like, care for, be fond of, dote on
v: annoy, upset, offend, hurt, wound, distress, insult, displease
v: upset, distress, trouble, disturb, worry, sadden, hurt, unsettle
v: comfort, console, reassure, soothe, cheer, calm, support, encourage
v: entertain, amuse, delight, divert, please, charm, captivate, cheer
v: interest, fascinate, intrigue, absorb, engage, attract, captivate, appeal to
v: bore, tire, weary, fatigue, wear out, exhaust
v: tire, exhaust, wear out, drain, weary, fatigue, sap
v: confuse, puzzle, bewilder, baffle, perplex, mystify, muddle, mix up
v: shock, appal, horrify, disgust, outrage, offend, sicken, stun
v: disgust, revolt, sicken, repel, nauseate, appal, offend
v: impress, strike, influence, move, affect, inspire, dazzle, stir
v: inspire, encourage, motivate, stimulate, stir, spur, energise, galvanise
v: attract, draw, pull, appeal to, lure, tempt, entice, charm
v: tempt, entice, lure, attract, seduce, coax, persuade, invite
v: trust, believe in, rely on, depend on, count on, have faith in, confide in
v: respect, admire, esteem, value, honour, revere, look up to, think highly of
v: admire, respect, esteem, praise, applaud, look up to, value, approve of
v: envy, covet, begrudge, resent, be jealous of
v: deserve, merit, earn, warrant, justify, be worthy of, rate
v: belong to, be owned by, be held by, be the property of
v: grab, seize, snatch, take hold of, catch, clutch
v: hand, pass, give, deliver, present, convey
v: lend, loan, advance, provide, give
v: borrow, take, use, hire, rent, cadge
v: rent, hire, lease, let, charter
v: earn, merit, deserve, win, achieve, attain, gain
v: gain, acquire, obtain, get, win, earn, achieve, secure
v: lose, misplace, mislay, drop, let slip
v: waste, squander, fritter away, misuse, throw away, lavish, dissipate
v: hurry, speed, fly, dash, run, race, scramble
v: wander, roam, stray, drift, meander, ramble, rove
v: sneak, creep, slip, steal, slink, tiptoe, sidle, skulk
v: hide, take cover, lie low, go into hiding, shelter
v: shake hands, greet, welcome, salute
v: kiss, peck, embrace, caress
v: hug, embrace, cuddle, hold, clasp, squeeze, enfold
v: care, look after, tend, mind, watch over, nurse, protect, attend to
v: look after, care for, tend, mind, protect, watch, supervise, nurse
v: feed, nourish, nurture, provide for, sustain, cater for
v: raise, bring up, rear, nurture, foster, educate, care for
v: grow, cultivate, plant, raise, produce, farm, breed
v: plant, sow, seed, bed, set, implant
v: gather, pick, harvest, collect, pluck, reap
v: fix, decide, settle, set, arrange, agree, establish
v: book, reserve, order, engage, arrange, schedule, organise
v: order, request, ask for, book, reserve, apply for, send for
v: order, command, instruct, direct, tell, require, bid, charge
v: wait for, await, expect, anticipate, look forward to
v: look forward to, anticipate, await, long for, count the days
v: remind, prompt, jog the memory, nudge, warn
v: promote, advertise, publicise, market, push, advance, boost, plug
v: advertise, publicise, promote, announce, market, broadcast, display
v: announce, declare, proclaim, report, state, reveal, publish, broadcast
v: report, describe, tell, announce, communicate, relate, document, record
v: record, register, note, document, log, write down, enter, chronicle
v: film, record, shoot, video, photograph, capture
v: photograph, snap, shoot, capture, take, film
v: broadcast, transmit, air, televise, show, relay, screen
v: publish, issue, print, release, bring out, produce, circulate
v: pronounce, say, articulate, sound, voice, utter, enunciate
v: spell, write, put into words
v: translate, interpret, render, convert, transcribe, decode, construe
v: interpret, explain, understand, read, construe, translate, decode
v: summarise, sum up, outline, condense, precis, recap, abstract
v: outline, sketch, summarise, draft, describe, rough out, trace
v: list, record, catalogue, enumerate, itemise, register, note, index
v: label, tag, mark, name, identify, categorise, classify, brand
v: mark, label, stamp, tag, brand, identify, sign
v: sign, autograph, initial, endorse, inscribe, put one's name to
v: stamp, mark, imprint, print, seal, label, brand
`;
