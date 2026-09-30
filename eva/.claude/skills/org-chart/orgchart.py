#!/usr/bin/env python3
"""
orgchart.py - draw an organisation chart as a single SVG file, in the house
design: a dark canvas, the org name in wide capitals top left, a boxed leader,
an optional row of leadership boxes, and one card per team (a light header band
with the team name and headcount, the leads, then the members on a tree line).

Part of the /org-chart skill (eva/.claude/skills/org-chart). The skill turns a
staffing spreadsheet into the JSON below; this script turns the JSON into the
SVG. Standard library only, Python 3.8+.

USAGE
-----
    python orgchart.py "H:\\Eva\\documents\\images\\Team Org Chart.json"
        -> writes "Team Org Chart.svg" beside it

    python orgchart.py chart.json -o "H:\\Eva\\documents\\images\\orgchart.svg"
    python orgchart.py chart.json --overwrite     replace an existing SVG
    python orgchart.py --sample sample.json       write a worked example JSON

An existing SVG is never replaced without --overwrite. Messages go to stderr;
exit code 0 = written, 1 = could not read/write a file, 2 = bad data.

INPUT (JSON, UTF-8)
-------------------
    {
      "title": "Halcyon Orbital",                 required - top left
      "subtitle": "Organisation chart",           optional (this default)
      "as_of": "2026-09-30",                      optional: ISO date (shown as
                                                  30 SEPT 2026), free text, or
                                                  "" for no date; default today
      "theme": "dark",                            optional: "dark" | "light"
      "count": "filled",                          optional: team headcount -
                                                  "filled" (leads + members,
                                                  vacancies excluded), "all"
                                                  (vacancies included), "none"
      "top": {"name": "Mara Keller", "role": "Director"},       required
      "leadership": [                             optional row under the top
        {"name": "Tomas Lindqvist", "role": "Technical Director"},
        {"name": "Aiko Tanaka", "role": "Technical Director"}
      ],
      "teams": [
        {
          "name": "Propulsion",
          "reports_to": "Tomas Lindqvist",        optional - see LAYOUT
          "leads": [
            {"name": "Erik Johansson", "role": "Team Lead"},
            {"name": "Ana Ribeiro", "role": "Tech Lead"}
          ],
          "members": [
            {"name": "Victor Hale", "role": "Turbomachinery Engineer"},
            {"role": "Test Technician", "vacant": true}
          ]
        }
      ]
    }

A person is {"name", "role"}, plus "vacant": true for an unfilled position
(the name may then be left out; it is drawn as VACANT, dimmed). Names, team
names and the labels are drawn in capitals; member roles as written.

LAYOUT
------
  - No "leadership": the teams hang straight off the top box.
  - "leadership" and NO team has "reports_to": the design's shared layout -
    the top box feeds every leadership box, and the leadership boxes join
    into one line that feeds every team.
  - "leadership" and EVERY team has "reports_to" naming one of them: each
    leadership box sits over its own teams and feeds only those (teams are
    regrouped in leadership order, keeping their order within a group).
    Mixing the two is refused, because it cannot be drawn without lines
    crossing.
Every card is the same width and height: the width grows to fit the longest
text (estimated - SVG cannot wrap or measure text), the height to the largest
team. Teams sit in one row, so a chart of more than about eight teams gets
wide; split it (one chart per leadership box) rather than shrinking it.

FONTS
-----
No font is embedded (the endpoint is offline, and an SVG shown through an
<img> tag cannot load one anyway). Text uses Segoe UI / Consolas, which every
Windows machine has, falling back to Helvetica/Arial and Courier. Widths are
estimated for Segoe UI with room to spare.

TESTING
-------
    python orgchart.py --sample sample.json
    python orgchart.py sample.json
    start sample.svg            (opens it in the default browser)

The sample reproduces the reference design: a director, two technical
directors and five teams.
"""

import argparse
import datetime
import json
import os
import sys
from xml.sax.saxutils import escape as _xml_escape

# ---------------------------------------------------------------------------
# Design constants. Measured from the reference design (drawn 2000px wide);
# change them here, not in the drawing code.
# ---------------------------------------------------------------------------
SANS = "'Segoe UI', 'Helvetica Neue', Helvetica, Arial, sans-serif"
MONO = "Consolas, 'Cascadia Mono', 'Courier New', monospace"

THEMES = {
    "dark": {
        "bg": "#0b0b0b", "text": "#f2f2f0", "label": "#8c8c8c",
        "role": "#a0a0a0", "rule": "#262626", "line": "#4d4d4d",
        "top_fill": "#161616", "top_stroke": "#cfcfcf",
        "lead_fill": "#121212", "lead_stroke": "#6a6a6a",
        "card_fill": "#111111", "card_stroke": "#3a3a3a",
        "band_fill": "#efefec", "band_text": "#111111", "band_count": "#3a3a3a",
        "divider": "#2c2c2c", "tree": "#555555", "vacant": "#6a6a6a",
    },
    "light": {
        "bg": "#ffffff", "text": "#141414", "label": "#6b6b6b",
        "role": "#555555", "rule": "#dcdcdc", "line": "#9a9a9a",
        "top_fill": "#f4f4f2", "top_stroke": "#222222",
        "lead_fill": "#fafafa", "lead_stroke": "#8a8a8a",
        "card_fill": "#fafaf9", "card_stroke": "#cfcfcf",
        "band_fill": "#141414", "band_text": "#f5f5f3", "band_count": "#cfcfcf",
        "divider": "#e2e2e2", "tree": "#b0b0b0", "vacant": "#9a9a9a",
    },
}

MARGIN = 64            # left/right page margin
HEADER_BASELINE = 77   # baseline of the title and the subtitle
RULE_Y = 113           # the thin line under the header
BOTTOM_MARGIN = 40

TITLE = dict(size=22, weight=700, spacing=7.0, font="sans")
SUBTITLE = dict(size=13, weight=400, spacing=3.8, font="mono")

TOP_Y = 177            # top of the leader's box
TOP_W, TOP_H = 435, 113
TOP_LABEL = dict(size=12, weight=400, spacing=3.5, font="mono")
TOP_NAME = dict(size=24, weight=600, spacing=2.6, font="sans")
TOP_LABEL_DY, TOP_NAME_DY = 43, 80

LEAD_W, LEAD_H = 384, 97          # leadership boxes
LEAD_GAP = 127                    # between leadership boxes
LEAD_LABEL = dict(size=12, weight=400, spacing=3.5, font="mono")
LEAD_NAME = dict(size=20, weight=600, spacing=2.4, font="sans")
LEAD_LABEL_DY, LEAD_NAME_DY = 38, 69

DROP_1 = 45            # top box bottom -> the bar feeding the next row
DROP_2 = 46            # that bar -> the row below
JOIN_1 = 41            # leadership bottom -> the joining bar (shared layout)
JOIN_2 = 41            # joining bar -> the bus over the teams
BUS_TO_CARD = 51       # bus -> card tops

CARD_W, CARD_GAP = 353, 29
BAND_H = 61
BAND_NAME = dict(size=17, weight=700, spacing=3.6, font="sans")
BAND_COUNT = dict(size=13, weight=400, spacing=1.5, font="mono")
BAND_BASELINE = 38
PAD_X = 26             # text inset inside a card
PAD_RIGHT = 24

ROLE_LABEL = dict(size=11, weight=400, spacing=3.0, font="mono")
ROLE_NAME = dict(size=17, weight=600, spacing=1.6, font="sans")
LEADS_FIRST = 36       # band bottom -> first lead label baseline
LEAD_PITCH = 67        # label to label
LEAD_NAME_OFFSET = 25  # label baseline -> name baseline
DIVIDER_OFFSET = 23    # last lead name baseline -> divider

MEMBER_NAME = dict(size=15, weight=500, spacing=1.0, font="sans")
MEMBER_ROLE = dict(size=14, weight=400, spacing=0.3, font="sans")
TREE_DX, TICK_DX, MEMBER_DX = 33, 48, 62
FIRST_TICK = 40        # divider -> first tick
MEMBER_PITCH = 59
MEMBER_NAME_DY = 6     # tick -> name baseline
MEMBER_ROLE_DY = 29    # tick -> role baseline
CARD_BOTTOM_PAD = 55   # last tick -> card bottom

STROKE = 1.5

# Australian month abbreviations, as the reference design writes the date.
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sept",
          "Oct", "Nov", "Dec"]

COUNT_MODES = ("filled", "all", "none")


class ChartError(Exception):
    """Bad input data; the message says which field and how to fix it."""


# ---------------------------------------------------------------------------
# Text width estimation. SVG has no text measurement outside a browser, so
# widths are estimated from per-character advances for Segoe UI (in em), on
# the generous side: a card a little wider than needed is fine, text running
# past its edge is not.
# ---------------------------------------------------------------------------
_NARROW_UPPER = {"I": 0.30, "J": 0.44, "L": 0.55, "F": 0.58, "E": 0.60,
                 "T": 0.62, "S": 0.62, "P": 0.64, "Z": 0.64}
_WIDE_UPPER = {"M": 0.90, "W": 0.98, "O": 0.76, "Q": 0.76, "G": 0.74,
               "C": 0.68, "D": 0.74, "H": 0.74, "N": 0.76, "U": 0.74}
_NARROW_LOWER = {"i": 0.26, "l": 0.26, "j": 0.28, "f": 0.36, "t": 0.38,
                 "r": 0.40, "s": 0.47}
_WIDE_LOWER = {"m": 0.86, "w": 0.78}
_PUNCT = {" ": 0.28, ".": 0.28, ",": 0.28, ":": 0.28, ";": 0.28, "'": 0.24,
          "-": 0.40, "(": 0.34, ")": 0.34, "/": 0.42, "&": 0.74,
          "\u00b7": 0.28}


def _char_em(ch):
    if ch in _PUNCT:
        return _PUNCT[ch]
    if ch.isdigit():
        return 0.58
    if ch.isupper():
        return _NARROW_UPPER.get(ch, _WIDE_UPPER.get(ch, 0.68))
    if ch.islower():
        return _NARROW_LOWER.get(ch, _WIDE_LOWER.get(ch, 0.55))
    return 0.72  # anything else (symbols, CJK): assume wide


def text_width(text, style):
    """Estimated rendered width in px of 'text' in 'style'."""
    if not text:
        return 0.0
    if style["font"] == "mono":
        em = 0.56 * len(text)
    else:
        em = sum(_char_em(ch) for ch in text)
        if style["weight"] >= 700:
            em *= 1.06
        elif style["weight"] >= 600:
            em *= 1.04
    # Browsers add letter-spacing after every character, the last included.
    return em * style["size"] + style["spacing"] * len(text)


# ---------------------------------------------------------------------------
# Input
# ---------------------------------------------------------------------------
def _str(value):
    return "" if value is None else str(value).strip()


def _person(raw, where, need_name=True):
    """Validate one person dict -> {'name', 'role', 'vacant'}."""
    if not isinstance(raw, dict):
        raise ChartError("{} must be an object like {{\"name\": ..., "
                         "\"role\": ...}}, got {!r}.".format(where, raw))
    vacant = raw.get("vacant") is True or _str(raw.get("vacant")).lower() in (
        "true", "yes", "1")
    name = _str(raw.get("name"))
    if name.lower() == "vacant":
        vacant, name = True, ""
    if need_name and not name and not vacant:
        raise ChartError("{} has no name. Give it one, or mark the position "
                         "\"vacant\": true.".format(where))
    return {"name": name, "role": _str(raw.get("role")), "vacant": vacant}


def format_as_of(value):
    """The date shown after 'AS OF': an ISO date becomes '30 SEPT 2026'."""
    if value is None:
        day = datetime.date.today()
    else:
        text = _str(value)
        if not text:
            return ""
        try:
            day = datetime.date.fromisoformat(text[:10]) if len(text) >= 10 else None
        except ValueError:
            day = None
        if day is None:
            return text.upper()
    return "{} {} {}".format(day.day, MONTHS[day.month - 1], day.year).upper()


def load_chart(data):
    """Validate the parsed JSON and return a normalised chart dict."""
    if not isinstance(data, dict):
        raise ChartError("The file must hold one JSON object (see the "
                         "docstring at the top of orgchart.py).")
    title = _str(data.get("title"))
    if not title:
        raise ChartError("\"title\" is required - the organisation's name, "
                         "shown top left.")
    theme = _str(data.get("theme")).lower() or "dark"
    if theme not in THEMES:
        raise ChartError("\"theme\" must be one of {} (got {!r}).".format(
            ", ".join(THEMES), data.get("theme")))
    count = _str(data.get("count")).lower() or "filled"
    if count not in COUNT_MODES:
        raise ChartError("\"count\" must be one of {} (got {!r}).".format(
            ", ".join(COUNT_MODES), data.get("count")))
    if "top" not in data:
        raise ChartError("\"top\" is required - the person the whole chart "
                         "reports to, e.g. {\"name\": \"Mara Keller\", "
                         "\"role\": \"Director\"}.")
    top = _person(data.get("top"), "\"top\"")

    leadership = data.get("leadership") or []
    if not isinstance(leadership, list):
        raise ChartError("\"leadership\" must be a list of people.")
    leadership = [_person(p, "leadership[{}]".format(i))
                  for i, p in enumerate(leadership)]
    lead_names = [p["name"] for p in leadership]
    named = [n for n in lead_names if n]
    if len(set(n.lower() for n in named)) != len(named):
        raise ChartError("Two leadership boxes have the same name, so a "
                         "team's \"reports_to\" could not tell them apart.")

    teams_raw = data.get("teams") or []
    if not isinstance(teams_raw, list):
        raise ChartError("\"teams\" must be a list.")
    teams = []
    for i, raw in enumerate(teams_raw):
        where = "teams[{}]".format(i)
        if not isinstance(raw, dict):
            raise ChartError("{} must be an object.".format(where))
        name = _str(raw.get("name"))
        if not name:
            raise ChartError("{} has no \"name\".".format(where))
        where = "team {!r}".format(name)
        leads = raw.get("leads") or []
        members = raw.get("members") or []
        if not isinstance(leads, list) or not isinstance(members, list):
            raise ChartError("{}: \"leads\" and \"members\" must be lists."
                             .format(where))
        teams.append({
            "name": name,
            "reports_to": _str(raw.get("reports_to")),
            "leads": [_person(p, "{} leads[{}]".format(where, j))
                      for j, p in enumerate(leads)],
            "members": [_person(p, "{} members[{}]".format(where, j))
                        for j, p in enumerate(members)],
        })

    # Which layout: see LAYOUT in the docstring.
    with_parent = [t for t in teams if t["reports_to"]]
    if with_parent and not leadership:
        raise ChartError(
            "Team {!r} has \"reports_to\" but there is no \"leadership\" row "
            "for it to report to. Remove \"reports_to\" (teams then hang off "
            "the top box) or add the leadership row.".format(
                with_parent[0]["name"]))
    grouped = bool(with_parent)
    if grouped:
        missing = [t["name"] for t in teams if not t["reports_to"]]
        if missing:
            raise ChartError(
                "Some teams have \"reports_to\" and some do not ({}). Either "
                "every team names the leadership box it reports to, or none "
                "does (they then report to the leadership row jointly)."
                .format(", ".join(repr(m) for m in missing)))
        lookup = {n.lower(): i for i, n in enumerate(lead_names) if n}
        for t in teams:
            idx = lookup.get(t["reports_to"].lower())
            if idx is None:
                raise ChartError(
                    "Team {!r} reports to {!r}, who is not in \"leadership\" "
                    "({}).".format(t["name"], t["reports_to"],
                                   ", ".join(repr(n) for n in named) or "none"))
            t["group"] = idx
        # Regroup in leadership order; sorted() is stable, so each group
        # keeps the order the teams were given in.
        teams = sorted(teams, key=lambda t: t["group"])

    return {
        "title": title,
        "subtitle": _str(data.get("subtitle")) or "Organisation chart",
        "as_of": format_as_of(data.get("as_of")),
        "theme": theme,
        "count": count,
        "top": top,
        "leadership": leadership,
        "teams": teams,
        "grouped": grouped,
    }


# ---------------------------------------------------------------------------
# Drawing
# ---------------------------------------------------------------------------
def _esc(text):
    return _xml_escape(text, {'"': "&quot;"})


def _num(value):
    """Coordinates to at most one decimal place, without a trailing .0."""
    value = round(value, 1)
    return str(int(value)) if value == int(value) else str(value)


class Svg:
    def __init__(self, colours):
        self.c = colours
        self.parts = []

    def rect(self, x, y, w, h, fill, stroke=None):
        extra = (' stroke="{}" stroke-width="{}"'.format(stroke, STROKE)
                 if stroke else "")
        self.parts.append('<rect x="{}" y="{}" width="{}" height="{}" '
                          'fill="{}"{}/>'.format(_num(x), _num(y), _num(w),
                                                 _num(h), fill, extra))

    def line(self, x1, y1, x2, y2, colour, dashed=False):
        dash = ' stroke-dasharray="3 3"' if dashed else ""
        self.parts.append('<line x1="{}" y1="{}" x2="{}" y2="{}" stroke="{}" '
                          'stroke-width="{}"{}/>'.format(
                              _num(x1), _num(y1), _num(x2), _num(y2), colour,
                              STROKE, dash))

    def polyline(self, points, colour):
        pts = " ".join("{},{}".format(_num(x), _num(y)) for x, y in points)
        self.parts.append('<polyline points="{}" fill="none" stroke="{}" '
                          'stroke-width="{}"/>'.format(pts, colour, STROKE))

    def text(self, x, y, text, style, colour, anchor="start"):
        if not text:
            return
        # Letter-spacing is added after every character, the last included,
        # so centred or right-aligned text would sit half a gap (or a whole
        # one) off. Shift it back by that much.
        if anchor == "middle":
            x += style["spacing"] / 2.0
        elif anchor == "end":
            x += style["spacing"]
        self.parts.append(
            '<text x="{}" y="{}" font-family="{}" font-size="{}" '
            'font-weight="{}" letter-spacing="{}" fill="{}"{}>{}</text>'.format(
                _num(x), _num(y), _esc(MONO if style["font"] == "mono" else SANS),
                style["size"], style["weight"], style["spacing"], colour,
                ' text-anchor="{}"'.format(anchor) if anchor != "start" else "",
                _esc(text)))


def _team_count(team, mode):
    people = team["leads"] + team["members"]
    if mode == "all":
        return len(people)
    return sum(1 for p in people if not p["vacant"])


def _display_name(person):
    return "VACANT" if person["vacant"] and not person["name"] else person["name"].upper()


def _leads_height(team):
    """Band bottom -> divider, for this team's leads."""
    n = len(team["leads"])
    if n == 0:
        return 0
    return LEADS_FIRST + (n - 1) * LEAD_PITCH + LEAD_NAME_OFFSET + DIVIDER_OFFSET


def measure(chart):
    """Card width, leadership/top box widths and card height for the chart."""
    count_w = text_width("00", BAND_COUNT)
    card_w = CARD_W
    for team in chart["teams"]:
        need = [PAD_X + text_width(team["name"].upper(), BAND_NAME) + 20
                + (count_w if chart["count"] != "none" else 0) + PAD_RIGHT]
        for p in team["leads"]:
            need.append(PAD_X + max(text_width(_display_name(p), ROLE_NAME),
                                    text_width(p["role"].upper(), ROLE_LABEL))
                        + PAD_RIGHT)
        for p in team["members"]:
            need.append(MEMBER_DX + max(text_width(_display_name(p), MEMBER_NAME),
                                        text_width(p["role"], MEMBER_ROLE))
                        + PAD_RIGHT)
        card_w = max([card_w] + need)

    lead_w = LEAD_W
    for p in chart["leadership"]:
        lead_w = max(lead_w, 48 + text_width(_display_name(p), LEAD_NAME),
                     48 + text_width(p["role"].upper(), LEAD_LABEL))
    top = chart["top"]
    top_w = max(TOP_W, 64 + text_width(_display_name(top), TOP_NAME),
                64 + text_width(top["role"].upper(), TOP_LABEL))

    leads_h = max([_leads_height(t) for t in chart["teams"]] or [0])
    most = max([len(t["members"]) for t in chart["teams"]] or [0])
    if most:
        body = leads_h + FIRST_TICK + (most - 1) * MEMBER_PITCH + CARD_BOTTOM_PAD
    else:
        body = max(leads_h, 40)
    return {
        "card_w": round(card_w), "lead_w": round(lead_w), "top_w": round(top_w),
        "leads_h": leads_h, "card_h": BAND_H + body,
        "widened": card_w > CARD_W,
    }


def layout_columns(chart, m):
    """
    x positions (left edges) of the cards and centres of the leadership
    boxes, relative to x=0 at the start of the content, plus content width.
    """
    teams, leaders = chart["teams"], chart["leadership"]
    card_w, lead_w = m["card_w"], m["lead_w"]
    card_x, lead_cx = [], []
    if chart["grouped"]:
        # One slot per leadership box: its teams side by side, the box
        # centred over them. A slot is never narrower than its box, and
        # slots are two card gaps apart so neighbouring groups read as groups.
        x = 0.0
        for g in range(len(leaders)):
            members = [i for i, t in enumerate(teams) if t["group"] == g]
            span = len(members) * card_w + max(len(members) - 1, 0) * CARD_GAP
            slot = max(span, lead_w)
            start = x + (slot - span) / 2.0
            for k, _i in enumerate(members):
                card_x.append(start + k * (card_w + CARD_GAP))
            lead_cx.append(x + slot / 2.0)
            x += slot + 2 * CARD_GAP
        width = x - 2 * CARD_GAP if leaders else 0.0
        return card_x, lead_cx, max(width, m["top_w"])

    cards_span = len(teams) * card_w + max(len(teams) - 1, 0) * CARD_GAP
    leads_span = len(leaders) * lead_w + max(len(leaders) - 1, 0) * LEAD_GAP
    width = max(cards_span, leads_span, m["top_w"])
    start = (width - cards_span) / 2.0
    card_x = [start + i * (card_w + CARD_GAP) for i in range(len(teams))]
    first = (width - leads_span) / 2.0 + lead_w / 2.0
    lead_cx = [first + i * (lead_w + LEAD_GAP) for i in range(len(leaders))]
    return card_x, lead_cx, width


def render(chart):
    """The chart as an SVG document (str), plus notes for the caller."""
    c = THEMES[chart["theme"]]
    m = measure(chart)
    card_x, lead_cx, content_w = layout_columns(chart, m)

    # Canvas width: the content, or the header if that is wider (a chart of
    # one team would otherwise squash the title into the subtitle).
    right_text = chart["subtitle"].upper()
    if chart["as_of"]:
        right_text += " \u00b7 AS OF " + chart["as_of"]
    header_w = (text_width(chart["title"].upper(), TITLE) + 80
                + text_width(right_text, SUBTITLE))
    inner_w = max(content_w, header_w)
    offset = MARGIN + (inner_w - content_w) / 2.0
    width = inner_w + 2 * MARGIN
    cx = offset + content_w / 2.0
    card_x = [offset + x for x in card_x]
    lead_cx = [offset + x for x in lead_cx]

    svg = Svg(c)
    teams, leaders = chart["teams"], chart["leadership"]

    # Vertical positions.
    top_bottom = TOP_Y + TOP_H
    if leaders:
        lead_top = top_bottom + DROP_1 + DROP_2
        lead_bottom = lead_top + LEAD_H
        if chart["grouped"]:
            bus_y = lead_bottom + JOIN_1 + JOIN_2
        else:
            join_y = lead_bottom + JOIN_1
            bus_y = join_y + JOIN_2
    else:
        bus_y = top_bottom + DROP_1
    card_top = bus_y + BUS_TO_CARD
    height = (card_top + m["card_h"] if teams else
              (lead_bottom if leaders else top_bottom)) + BOTTOM_MARGIN

    # Header.
    svg.text(MARGIN, HEADER_BASELINE, chart["title"].upper(), TITLE, c["text"])
    svg.text(width - MARGIN, HEADER_BASELINE, right_text, SUBTITLE, c["label"],
             anchor="end")
    svg.line(MARGIN, RULE_Y, width - MARGIN, RULE_Y, c["rule"])

    # Connectors first, so the boxes are drawn over their ends.
    card_cx = [x + m["card_w"] / 2.0 for x in card_x]

    def bus(y, drops, drop_to, reach=()):
        """
        A horizontal bar at y with a line down from it to drop_to at each x in
        'drops'. 'reach' holds x positions the bar must also span without
        dropping - where the line from above lands on it.
        """
        span = list(drops) + list(reach)
        if len(span) > 1 and max(span) - min(span) > 0.05:
            svg.line(min(span), y, max(span), y, c["line"])
        for x in drops:
            svg.line(x, y, x, drop_to, c["line"])

    if leaders:
        mid_y = top_bottom + DROP_1
        svg.line(cx, top_bottom, cx, mid_y, c["line"])
        bus(mid_y, lead_cx, lead_top, reach=[cx])
        if chart["grouped"]:
            for g, lx in enumerate(lead_cx):
                xs = [card_cx[i] for i, t in enumerate(teams) if t["group"] == g]
                if not xs:
                    continue
                svg.line(lx, lead_bottom, lx, bus_y, c["line"])
                bus(bus_y, xs, card_top, reach=[lx])
        elif teams:
            for lx in lead_cx:
                svg.line(lx, lead_bottom, lx, join_y, c["line"])
            bus(join_y, [], 0, reach=lead_cx + [cx])
            svg.line(cx, join_y, cx, bus_y, c["line"])
            bus(bus_y, card_cx, card_top, reach=[cx])
    elif teams:
        svg.line(cx, top_bottom, cx, bus_y, c["line"])
        bus(bus_y, card_cx, card_top, reach=[cx])

    # Top box.
    top = chart["top"]
    tx = cx - m["top_w"] / 2.0
    svg.rect(tx, TOP_Y, m["top_w"], TOP_H, c["top_fill"], c["top_stroke"])
    svg.text(cx, TOP_Y + TOP_LABEL_DY, top["role"].upper(), TOP_LABEL,
             c["label"], anchor="middle")
    svg.text(cx, TOP_Y + TOP_NAME_DY, _display_name(top), TOP_NAME,
             c["vacant"] if top["vacant"] else c["text"], anchor="middle")

    # Leadership boxes.
    for p, lx in zip(leaders, lead_cx):
        svg.rect(lx - m["lead_w"] / 2.0, lead_top, m["lead_w"], LEAD_H,
                 c["lead_fill"], c["lead_stroke"])
        svg.text(lx, lead_top + LEAD_LABEL_DY, p["role"].upper(), LEAD_LABEL,
                 c["label"], anchor="middle")
        svg.text(lx, lead_top + LEAD_NAME_DY, _display_name(p), LEAD_NAME,
                 c["vacant"] if p["vacant"] else c["text"], anchor="middle")

    # Team cards.
    for team, x in zip(teams, card_x):
        w, h = m["card_w"], m["card_h"]
        svg.rect(x, card_top, w, h, c["card_fill"], c["card_stroke"])
        svg.rect(x, card_top, w, BAND_H, c["band_fill"])
        svg.text(x + PAD_X, card_top + BAND_BASELINE, team["name"].upper(),
                 BAND_NAME, c["band_text"])
        if chart["count"] != "none":
            svg.text(x + w - PAD_RIGHT, card_top + BAND_BASELINE,
                     "{:02d}".format(_team_count(team, chart["count"])),
                     BAND_COUNT, c["band_count"], anchor="end")

        body_top = card_top + BAND_H
        for j, p in enumerate(team["leads"]):
            label_y = body_top + LEADS_FIRST + j * LEAD_PITCH
            svg.text(x + PAD_X, label_y, p["role"].upper(), ROLE_LABEL, c["label"])
            svg.text(x + PAD_X, label_y + LEAD_NAME_OFFSET, _display_name(p),
                     ROLE_NAME, c["vacant"] if p["vacant"] else c["text"])
        # Every card's divider sits at the same height (the tallest leads
        # block), so members line up across the row.
        divider_y = body_top + m["leads_h"]
        if team["leads"] and team["members"]:
            svg.line(x, divider_y, x + w, divider_y, c["divider"])

        if team["members"]:
            ticks = [divider_y + FIRST_TICK + k * MEMBER_PITCH
                     for k in range(len(team["members"]))]
            svg.line(x + TREE_DX, divider_y, x + TREE_DX, ticks[-1], c["tree"])
            for p, ty in zip(team["members"], ticks):
                svg.line(x + TREE_DX, ty, x + TICK_DX, ty, c["tree"],
                         dashed=p["vacant"])
                svg.text(x + MEMBER_DX, ty + MEMBER_NAME_DY, _display_name(p),
                         MEMBER_NAME, c["vacant"] if p["vacant"] else c["text"])
                svg.text(x + MEMBER_DX, ty + MEMBER_ROLE_DY, p["role"],
                         MEMBER_ROLE, c["role"])

    people = 1 + len(leaders) + sum(len(t["leads"]) + len(t["members"])
                                    for t in teams)
    label = "{} organisation chart".format(chart["title"])
    head = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" '
        'viewBox="0 0 {w} {h}" role="img" aria-label="{label}">\n'
        '<title>{label}</title>\n'
        '<desc>{desc}</desc>\n'
        '<rect width="100%" height="100%" fill="{bg}"/>\n'
    ).format(w=_num(width), h=_num(height), label=_esc(label), bg=c["bg"],
             desc=_esc("{} people in {} team{}{}".format(
                 people, len(teams), "" if len(teams) == 1 else "s",
                 ", as of " + chart["as_of"].title() if chart["as_of"] else "")))
    doc = head + "\n".join(svg.parts) + "\n</svg>\n"

    notes = {
        "width": round(width), "height": round(height), "people": people,
        "teams": len(teams),
        "vacant": sum(1 for p in [top] + leaders
                      + [q for t in teams for q in t["leads"] + t["members"]]
                      if p["vacant"]),
        "card_w": m["card_w"], "widened": m["widened"],
    }
    return doc, notes


# ---------------------------------------------------------------------------
# Sample data: the reference design.
# ---------------------------------------------------------------------------
def _p(name, role):
    return {"name": name, "role": role}


SAMPLE = {
    "title": "Halcyon Orbital",
    "subtitle": "Organisation chart",
    "as_of": "2026-09-30",
    "theme": "dark",
    "count": "filled",
    "top": _p("Mara Keller", "Director"),
    "leadership": [_p("Tomas Lindqvist", "Technical Director"),
                   _p("Aiko Tanaka", "Technical Director")],
    "teams": [
        {"name": "Propulsion",
         "leads": [_p("Erik Johansson", "Team Lead"), _p("Ana Ribeiro", "Tech Lead")],
         "members": [_p("Victor Hale", "Turbomachinery Engineer"),
                     _p("Mei Lin", "Combustion Engineer"),
                     _p("Jonah Weber", "Propellant Systems Engineer"),
                     _p("Liam Carter", "Test Technician")]},
        {"name": "Structures",
         "leads": [_p("Nadia Petrova", "Team Lead"), _p("Owen Hughes", "Tech Lead")],
         "members": [_p("Isabel Costa", "Stress Analyst"),
                     _p("Kenji Mori", "Materials Engineer"),
                     _p("Sofia Nguyen", "Design Engineer")]},
        {"name": "Avionics",
         "leads": [_p("Marcus Reid", "Team Lead"), _p("Leah Friedman", "Tech Lead")],
         "members": [_p("Samuel Park", "Hardware Engineer"),
                     _p("Beatriz Alves", "Firmware Engineer"),
                     _p("Noah Bennett", "Harness Engineer"),
                     _p("Grace Adeyemi", "Test Engineer"),
                     _p("Rafael Moreno", "RF Engineer")]},
        {"name": "Guidance & Control",
         "leads": [_p("Yusuf Demir", "Team Lead"), _p("Hannah Brooks", "Tech Lead")],
         "members": [_p("Daniel Osei", "GNC Engineer"),
                     _p("Priya Raman", "Flight Software Engineer"),
                     _p("Clara Weiss", "Simulation Engineer")]},
        {"name": "Flight Software",
         "leads": [_p("Ana Silva", "Team Lead"), _p("Ben Ortiz", "Tech Lead")],
         "members": [_p("Chloe Martin", "Software Engineer"),
                     _p("Ravi Patel", "Software Engineer"),
                     _p("Elena Rossi", "Test Automation Engineer"),
                     _p("Jack Dunn", "Ground Systems Engineer")]},
    ],
}


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------
def err(message):
    print("orgchart: " + message, file=sys.stderr)


def write_file(path, text, overwrite):
    """Write text to path via a temporary file; returns an exit code."""
    if os.path.exists(path) and not overwrite:
        err("{} already exists, so nothing was written. Re-run with "
            "--overwrite to replace it, or -o for another name.".format(path))
        return 1
    folder = os.path.dirname(os.path.abspath(path))
    if not os.path.isdir(folder):
        err("The folder {} does not exist. Create it (or copy the eva\\ "
            "folder to H:\\Eva) and try again.".format(folder))
        return 1
    partial = path + ".part"
    try:
        with open(partial, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        os.replace(partial, path)
    except OSError as e:
        try:
            os.remove(partial)
        except OSError:
            pass
        err("Could not write {} ({}). If it is open in another program, "
            "close it and try again.".format(path, e))
        return 1
    return 0


def build_parser():
    parser = argparse.ArgumentParser(
        description="Draw an organisation chart as an SVG from a JSON file "
                    "(see the docstring at the top of this file for the format).")
    parser.add_argument("input", nargs="?",
                        help="the chart's JSON file")
    parser.add_argument("-o", "--output",
                        help="the SVG to write (default: the input's name "
                             "with .svg)")
    parser.add_argument("--overwrite", action="store_true",
                        help="replace the output file if it already exists")
    parser.add_argument("--sample", metavar="PATH",
                        help="write the worked example JSON to PATH and exit")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)

    if args.sample:
        code = write_file(args.sample, json.dumps(SAMPLE, indent=2,
                                                  ensure_ascii=False) + "\n",
                          args.overwrite)
        if code == 0:
            err("wrote the sample chart to {}".format(args.sample))
        return code

    if not args.input:
        build_parser().print_usage(sys.stderr)
        err("give the chart's JSON file (or --sample PATH for an example).")
        return 2
    try:
        # utf-8-sig: a file saved by Notepad may start with a byte-order mark.
        with open(args.input, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
    except FileNotFoundError:
        err("No file at {}.".format(args.input))
        return 1
    except OSError as e:
        err("Could not read {} ({}).".format(args.input, e))
        return 1
    except ValueError as e:
        err("{} is not valid JSON: {}".format(args.input, e))
        return 2

    try:
        chart = load_chart(data)
    except ChartError as e:
        err(str(e))
        return 2
    doc, notes = render(chart)

    output = args.output or os.path.splitext(args.input)[0] + ".svg"
    code = write_file(output, doc, args.overwrite)
    if code:
        return code
    err("wrote {} ({} x {} px): {} people, {} team{}{}{}".format(
        output, notes["width"], notes["height"], notes["people"],
        notes["teams"], "" if notes["teams"] == 1 else "s",
        ", {} vacant".format(notes["vacant"]) if notes["vacant"] else "",
        ", {} layout".format("grouped" if chart["grouped"] else "shared")
        if chart["leadership"] else ""))
    if notes["widened"]:
        err("note: cards widened to {} px to fit the longest name or role "
            "(design width {}).".format(notes["card_w"], CARD_W))
    if notes["teams"] > 8:
        err("note: {} teams in one row makes a very wide chart; consider one "
            "chart per leadership box.".format(notes["teams"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
