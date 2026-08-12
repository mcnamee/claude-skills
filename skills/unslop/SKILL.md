---
name: unslop
description: Strip AI-slop markers from a piece of writing — padding phrases, tell-tale vocabulary, and the stock LLM sentence and list shapes — without changing the meaning or the author's voice. Use ONLY when explicitly asked to de-slop, unslop, "make this sound less like AI", or clean up AI-sounding text. Do not apply it unasked to your own drafts or to text the user merely shared.
---

# unslop

Remove the markers that make writing read as machine-generated, and nothing
else. The output must say exactly what the input said, in the same voice.

This is a **subtractive** edit. You are not improving the writing, tightening
the argument, restructuring for clarity, or making it punchier. Those are
different jobs — if the user wants them, they will ask.

## 1. Get the input

In order of preference:

1. Text passed with the command, or pasted in the same message.
2. A file path given with the command — read it. Edit in place only if the
   user asked for that; otherwise return the cleaned text.
3. If the command arrives bare, use the most recent substantial piece of
   writing in the conversation — and say which one you took, in one line,
   before the output.

If there is nothing to work on, ask for the text. Do not invent an example.

## 2. Read for voice before you cut

Before changing anything, read a few paragraphs and note the author's habits:
sentence length, contractions, formality, British vs American spelling,
whether they use em dashes, first person, humour. **These are the things you
must preserve.** A de-slopped text that no longer sounds like its author is a
failed edit, even if every marker is gone.

Some authors genuinely write with em dashes, or genuinely say "delve". If a
habit is consistent across the whole piece, it is voice, not slop.

## 3. Tier 1 — delete outright

Padding that carries no information. Cut the phrase, keep the sentence.

| Cut | Note |
|---|---|
| "It's important to note that…", "It's worth noting that…", "It should be noted that…" | The sentence that follows is the content |
| "In today's fast-paced world", "In an era of…", "In the ever-evolving landscape of…" | Openers that say nothing |
| "Let's dive in", "Let's explore", "Let's unpack this" | |
| "Great question!", "Absolutely!", "You're absolutely right!" | Sycophantic openers |
| "I hope this helps!", "Feel free to reach out" | Sign-offs the author did not write |
| "As an AI language model…", "I don't have personal opinions, but…" | |
| "At the end of the day", "When all is said and done", "Needless to say" | |
| "In conclusion," / "In summary," | Cut the label; cut the whole paragraph if it only restates what came before and adds nothing |
| "The key benefits include:", "Here are the key takeaways:", "This offers several advantages:" | List stems that announce a list instead of saying anything — see Tier 3, *the label-colon list* |

A closing paragraph that recaps the piece with no new information is itself
slop. Delete it — unless it lands a conclusion the body did not state.

## 4. Tier 2 — swap for the plain word

Vocabulary tells. Replace with the ordinary word, or cut if it adds nothing.

| Marker | Plain |
|---|---|
| delve into | look at, examine, go into |
| leverage *(verb)* | use |
| utilize | use |
| facilitate | help, make possible |
| foster | encourage, build |
| harness | use |
| embark on | start |
| navigate *(metaphorical)* | handle, deal with, work through |
| underscore, highlight *(as "shows")* | show, stress |
| a testament to | shows, proof of |
| robust | strong, reliable, thorough |
| seamless(ly) | smooth, without interruption |
| myriad, plethora, a wide range of | many |
| crucial, pivotal, vital | important — or cut |
| cutting-edge, state-of-the-art | new, current — or cut |
| realm, landscape, tapestry, ecosystem *(metaphorical)* | the plain noun: field, market, set of tools |
| resonate with | matter to, appeal to |
| in order to | to |
| game-changer, revolutionary | say what actually changed |

Two cautions. **Terms of art are not slop:** "robust" in statistics,
"leverage" in finance, "navigate" about actual navigation, "harness" about
actual harnesses. And **do not swap a word the author uses consistently** —
that is voice.

## 5. Tier 3 — restructure the sentence or the list

The stock LLM shapes. These are the strongest tells, and the ones worth the
most care, because fixing them means rewriting a clause.

**The label-colon list.** A stem that announces a list, then bullets of the
shape `**Abstract Noun Phrase:** fragment` — or
`<strong>Abstract Noun Phrase:</strong> fragment`, the same pattern in HTML
clothing:

```
The key benefits include:
- **Better alignment:** Your staff will thrive.
- **Cross organisational synergy:** Making use of each other's strengths.
- **De-duplication of effort:** Ensuring best athlete approach.
```

Three markers stacked: a stem that praises rather than introduces, labels
coined so that there is something to bold, and bodies that are fragments
instead of sentences. Usually exactly three bullets.

*The test — cover the label and see what is lost.* If it is a term the reader
would scan or search for — a parameter, an option, a product, a person, a
date, a numbered step — it is load-bearing and the list is fine. If it is a
nominalisation of the sentence sitting next to it ("Better alignment:" in
front of "Your staff will thrive"), it carries nothing, and the colon is
standing in for a sentence the author never wrote.

Two repairs, in order of preference:

1. **Drop the labels, keep the bullets.** The claims survive; the packaging
   goes. "**De-duplication of effort:** Ensuring best athlete approach." →
   "Effort is not duplicated."
2. **Fold the bullets into prose** when they are one sentence that has been
   chopped into three — the Tier 4 bullet-list rule, applied with more
   confidence here, because the labels have already shown you the bullets are
   not really a list.

Then the stem. Cut it if the sentence before it already introduces the list.
If the list does need a lead-in, let it name what follows rather than
advertise it, and strip the empty intensifier: "The key benefits include:" →
"Benefits:". Never keep "key", "several", "a number of", "a few" in that
position — they are Tier 1 padding wearing a colon.

Hold the line on content: **the labels are packaging, not claims.** Do not
invent a specific to replace a vague label, and do not drop a bullet because
its claim is thin — "Your staff will thrive" comes out exactly as vague as it
went in.

**Negation-antithesis.** "It's not just a text editor — it's a way of
thinking." → State the claim: "It's a way of thinking." Keep the contrast
only when both halves carry information the reader needs.

**Participial tails.** ", ensuring reliability", ", allowing teams to move
faster", ", making it ideal for large documents". Either make it a real
clause ("so it stays reliable") or cut it, because it usually just restates
the main clause in different words.

**Rhetorical fragment questions.** "The result? Faster builds." "The catch?
It's Linux-only." → "The result is faster builds." / "It's Linux-only,
though."

**Empty triads.** "clear, concise, and compelling" — keep the one or two
adjectives that are true and load-bearing, drop the one added for rhythm. The
same instinct sets lists to three items by default. Treat that count as a
prompt to check whether each item carries its own claim, not as something to
fix: you can neither add a claim nor drop one.

**Symmetrical scaffolding.** "Firstly… Secondly… Lastly…" where the order
carries no meaning. Drop the ordinals, or make it a real list.

**Throat-clearing openers.** "Here's the thing:", "Let's be clear:", "Make no
mistake:" → delete the opener, keep the sentence.

**Stacked hedges.** "may potentially be somewhat", "it could arguably be said
that" → one hedge at most. Keep hedges that reflect the author's genuine
uncertainty; cut the ones that are just softening.

**Both-sides non-endings.** "While X has benefits, it also has drawbacks.
Ultimately, it depends." Keep the specifics, cut the empty framing — but
**never invent a conclusion the author did not reach.** If there is no view
in the source, there is none in the output.

## 6. Tier 4 — signals, not targets

These correlate with AI writing but are also just… writing. Intervene only
when they are clearly mechanical, and prefer varying to purging.

- **Em dashes.** Legitimate punctuation. Act only when several in one short
  paragraph do the same job, and then vary the punctuation rather than
  removing the pauses.
- **Bullet lists.** Fine when the content is genuinely a list. Convert to
  prose only when each bullet is a fragment of a single argument that has
  been chopped up — and see Tier 3, *the label-colon list*, for the case where
  the bullets themselves tell you which one this is.
- **Bold.** Reduce when whole sentences are bold, or when every other phrase
  is emphasised. Leave deliberate key-term bolding. Bold that opens a bullet
  and ends in a colon is the Tier 3 pattern, not this one.
- **Raw HTML emphasis.** `<strong>`, `<b>`, `<em>` sitting in text that is
  otherwise Markdown or plain is a machine artefact: match the document's own
  convention (`**bold**`) or, if the emphasis was only there to prop up a
  label you have just removed, drop it. Leave the tags alone in a file that is
  genuinely HTML.
- **Emoji in headings.** Cut if the rest of the document has none. Leave if
  the document is consistently informal.
- **Headings and structure.** Leave alone. Over-sectioning is a slop signal
  but restructuring a document is not a subtractive edit.
- **"However", "Moreover", "Furthermore", "Additionally".** Judge each: if it
  states a real relationship, keep it. "Moreover"/"Furthermore" are usually
  deletable or replaceable with "And"/"Also"; "However" is often earning its
  place.

## 7. Never touch

- Facts, numbers, dates, names, units, citations, links
- **Quotations** — verbatim, even when the quoted text is itself sloppy
- Code, commands, config, file paths, identifiers, log output, error strings
- The author's opinions, jokes, asides, and genuine hedges
- Spelling conventions (British vs American), Oxford comma habit, capitalisation style
- Anything inside a language you cannot confidently judge

## 8. The ambiguity rule

**When you cannot tell whether something is slop or the author's voice, leave
it.** One surviving marker costs far less than a sentence that no longer
sounds like the person who wrote it.

If the text is already clean, say so and return it unchanged. Do not
manufacture edits to look useful.

## 9. Output

Return the cleaned text in the same format as the input — Markdown stays
Markdown, plain text stays plain. No preamble, no "here's the cleaned
version", no commentary wrapped around it.

If the user asks what changed, follow the text with a short table:

| Before | After | Why |
|---|---|---|
| It's important to note that the API is rate-limited | The API is rate-limited | padding |
| leverage the cache | use the cache | vocabulary |
| It's not just faster — it's cheaper | It's faster and cheaper | stock construction |
| **De-duplication of effort:** Ensuring best athlete approach. | Effort is not duplicated | label-colon list |

## 10. Check before you return

1. **Every claim in the input is still in the output, and no claim is new.**
2. Numbers, names, quotes and code are byte-identical.
3. The word count went down or stayed flat. If it went up, you rewrote
   instead of removing — start again.
4. Read the opening and closing aloud. Same person?
5. **If you changed more than about a third of the words, stop and justify
   each edit against a tier above.** Genuinely slop-dense text can warrant
   that much change — a paragraph that is mostly padding will lose most of
   its words. But that ratio is also what an unasked-for rewrite looks like,
   so any edit you cannot point at a tier for gets reverted.
