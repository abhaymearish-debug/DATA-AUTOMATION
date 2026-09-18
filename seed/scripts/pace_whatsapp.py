"""
Daily pace — WhatsApp message pack.

One number per executive, and it climbs when he falls behind. That is the
whole message. Blocks are ~10 lines on purpose: anything that needs scrolling
on a phone gets skimmed, and a skimmed target is not a target.

WhatsApp markup: *bold*, _italic_.
"""

from __future__ import annotations

CUT = "─" * 28


def _cs(v: float) -> str:
    return f"{v:,.0f}"


def _yesterday(b, as_of):
    for r in b["rows"]:
        if r["day"] == as_of.day:
            return r
    return None


def bond_block(plan, bond, targets, staff=None) -> str:
    b = plan["bonds"][bond]
    as_of = plan["as_of"]
    nxt = b["next_date"]
    y = _yesterday(b, as_of)
    behind = b["gap"] < 0

    L = [f"*{bond}*  ·  {nxt:%d %b}" if nxt else f"*{bond}*"]
    if staff and staff.upper() != "VACANT":
        L.append(f"_{staff.title()}_")
    L.append("")

    if y and y["dry"]:
        L.append(f"⚪ *{as_of:%d %b}:*  dry day")
    elif y and y["actual"] is not None:
        mark = "✅" if y["status"] == "hit" else "🔻"
        L.append(f"{mark} *{as_of:%d %b}:*  {_cs(y['actual'])} cs "
                 f"_(needed {_cs(y['target_day'])})_")

    L.append(f"🎯 *Shop target:*  {_cs(b['target_month'])} cs   "
             f"_({_cs(b['per_day'])} cs a day)_")
    L.append(f"📦 *Shop sales:*  {_cs(b['mtd_actual'])} cs  ·  "
                 f"{(b['mtd_actual']/b['mtd_expected']) if b['mtd_expected'] else 0:.0%} "
                 f"of where you should be")
    if behind:
        L.append(f"📉 Should be at {_cs(b['mtd_expected'])} — "
                 f"*behind by {_cs(-b['gap'])} cs*")
    else:
        L.append(f"📈 Should be at {_cs(b['mtd_expected'])} — "
                 f"*ahead by {_cs(b['gap'])} cs* ✅")
    if b["streak"] >= 3:
        L.append(f"🔥 *{b['streak']} days in a row* on target")
    L.append(f"🏆 Rank *{b['rank']} of 15*")
    L.append("")
    L.append(f"👉 *TODAY: {_cs(b['needed_per_day'])} cs*")
    L.append(f"_{b['days_left']} days left. Every day you fall short, "
             f"this number goes up._" if behind else
             f"_{b['days_left']} days left. Stay on it._")

    shops = targets.get(bond, [])
    if shops:
        drops = sorted((s for s in shops if s["delta"] < -1.0),
                       key=lambda s: s["delta"])[:3]
        L.append("")
        L.append("*Top shops — cases this month:*")
        for r in shops[:3]:
            L.append(f"• *{r['shop']}* — {r['now']:,.0f} cs "
                     f"_(same days last month {r['last']:,.0f})_")
        if drops:
            L.append("")
            L.append("*Slipped vs last month:*")
            for r in drops:
                L.append(f"• *{r['shop']}* — {r['now']:,.0f} cs, "
                         f"down {abs(r['delta']):,.0f}")
    return "\n".join(L)


def cluster_block(plan, cluster, bonds) -> str:
    rows = sorted((plan["bonds"][b] for b in bonds if b in plan["bonds"]),
                  key=lambda x: -x["pct_done"])
    nxt = rows[0]["next_date"] if rows else None
    ta = sum(r["mtd_actual"] for r in rows)
    tt = sum(r["target_month"] for r in rows)
    tn = sum(r["needed_per_day"] for r in rows)

    L = [f"*CLUSTER {cluster} — TABLE*  ·  {nxt:%d %b}" if nxt
         else f"*CLUSTER {cluster} — TABLE*", ""]
    for i, r in enumerate(rows, 1):
        icon = "🟢" if r["gap"] >= 0 else ("🟡" if r["pct_pace"] >= 0.85 else "🔴")
        # a medal on a bond running 27% reads as praise for missing
        medal = ({1: "🥇", 2: "🥈", 3: "🥉"}.get(i, f"{i}.")
                 if r["gap"] >= 0 else f"{i}.")
        L.append(f"{medal} *{r['bond']}*  {r['pct_done']:.0%} {icon}  "
                 f"_{_cs(r['mtd_actual'])}/{_cs(r['target_month'])} cs_  "
                 f"→ *{_cs(r['needed_per_day'])}/day*")
    L.append("")
    L.append(f"Cluster: *{_cs(ta)} / {_cs(tt)} cs* · "
             f"{(ta / tt if tt else 0):.0%} done")
    L.append(f"👉 *Needs {_cs(tn)} cs a day to finish*")
    return "\n".join(L)


def network_block(plan) -> str:
    n = plan["network"]
    nxt = next((b["next_date"] for b in plan["bonds"].values()
                if b["next_date"]), None)
    L = [f"*K.S. DISTILLERY — STATE*  ·  {nxt:%d %b}" if nxt
         else "*K.S. DISTILLERY — STATE*", ""]
    L.append(f"🎯 Target: *{_cs(n['target_month'])} cs*  "
             f"_({_cs(n['per_day'])} a day)_")
    L.append(f"📦 Sold: *{_cs(n['mtd_actual'])} cs*  ·  {n['pct_done']:.0%} done")
    L.append(f"📉 Behind by *{_cs(-n['gap'])} cs*" if n["gap"] < 0
             else f"📈 Ahead by *{_cs(n['gap'])} cs*")
    L.append("")
    L.append(f"👉 *{_cs(n['needed_per_day'])} cs a day* for the last "
             f"{n['days_left']} days")
    L.append("")
    ranked = [plan["bonds"][b] for b in plan["order"]]
    L.append("*Front:*  " + "   ".join(f"{r['bond'].title()} {r['pct_done']:.0%}"
                                       for r in ranked[:3]))
    L.append("*Back:*  " + "   ".join(f"{r['bond'].title()} {r['pct_done']:.0%}"
                                      for r in ranked[-3:]))
    return "\n".join(L)


def build_pack(plan, targets, master, clusters, asms=None) -> str:
    """
    The pack is cut for a CASCADE, not for one person sending 15 messages.

    Abhay pastes three things -- one section per ASM. Each section is
    self-contained: the cluster league table to drop in the cluster group,
    then that cluster's executive blocks for the ASM to forward on. Three
    actions for Abhay, five or six for each ASM, and the ASM ends up inside
    the accountability loop rather than beside it.
    """
    asms = asms or {}
    staff = {}
    for info in master.values():
        if info["status"].upper() == "CLOSED":
            continue
        s = (info["staff"] or "").strip()
        if info["bond"] and s and s.upper() != "VACANT":
            staff.setdefault(info["bond"], s)

    nxt = next((b["next_date"] for b in plan["bonds"].values()
                if b["next_date"]), None)
    rank = {b: i for i, b in enumerate(plan["order"])}

    out = [
        "\n".join(x for x in [
            f"DAILY PACE PACK — {plan['month']} {plan['year']}",
            f"For {nxt:%A %d %B %Y}" if nxt else "",
            f"Sales through {plan['as_of']:%d %b}.",
            "",
            f"HOW TO SEND — {len(clusters)} sections below, one per ASM.",
            "In each: paste the TABLE into that cluster's group, then forward "
            "each executive his own block.",
            "Daily target = monthly target ÷ selling days. "
            "Fall short and the number goes up.",
        ] if x),
        CUT,
        "*FOR ABHAY — do not forward*\n\n" + network_block(plan),
    ]

    for i, c in enumerate(sorted(clusters), 1):
        bonds = [b for b in clusters[c] if b in plan["bonds"]]
        bonds.sort(key=lambda b: rank.get(b, 99))
        asm = asms.get(c)
        out.append("═" * 28)
        out.append(f"SECTION {i} of {len(clusters)}  →  "
                   f"{(asm or f'CLUSTER {c} ASM').upper()}  ·  CLUSTER {c}\n"
                   f"Paste the table in the group, then forward "
                   f"{len(bonds)} blocks.")
        out.append(CUT)
        out.append(cluster_block(plan, c, bonds))
        for bond in bonds:
            out.append(CUT)
            out.append(bond_block(plan, bond, targets, staff.get(bond)))
        out.append("")
    return "\n\n".join(out)
