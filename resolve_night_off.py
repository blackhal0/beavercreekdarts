#!/usr/bin/env python3
"""
resolve_night_off.py — solves the "Team Nights Off" requests in
Dart_League_Schedule.xlsx and writes the resulting 15-week schedule back
into the workbook.

WHAT IT DOES
------------
Reads every request in the 'Team Nights Off' tab (Team letter + Night #)
and, starting fresh each time from the ORIGINAL optimized rotation stored
in the hidden 'Base_Rotation' sheet, finds a schedule that gets every team
off on every night they asked for -- while guaranteeing:
  - every pair of teams still plays each other exactly twice all season
  - every team still gets exactly 5 byes
For each request it tries, in order:
  1. Already on bye that night -> nothing to do.
  2. Self-swap -- trade this team's whole week with a DIFFERENT week where
     it's already on bye (only works if that night hasn't already been
     locked in by an earlier, different team's request).
  3. Borrow -- if the night is already locked for another team, find the
     team currently sharing that bye slot who ISN'T locked there, and swap
     just those two teams' roles between this night and one other night,
     chosen so nobody's total games-per-opponent or bye count changes.
If no valid swap exists for a request, it's reported as a genuine conflict
(this happens only when the schedule's fixed bye-partner structure makes a
combination of requests impossible -- e.g. a 3rd team wanting a night that
two locked teams already occupy).

USAGE
-----
    python3 resolve_night_off.py Dart_League_Schedule.xlsx

Run this every time you add, change, or remove a row in the Team Nights
Off tab. It edits the file in place (back up first if you want to be
safe) and prints a log of exactly what it did.

Requires: openpyxl, and LibreOffice for the recalculation step
(the script calls the xlsx skill's recalc helper if it's available on
PATH as `soffice`; otherwise just re-open the file in Excel once to
refresh the Schedule/Standings tabs).
"""
import sys
import subprocess
import shutil
import openpyxl

ROLES = ['B1T1', 'B1T2', 'B2T1', 'B2T2', 'ByeT1', 'ByeT2']


def load_base_rotation(wb):
    ws = wb['Base_Rotation']
    grid = {}
    for row in range(2, 17):
        week = ws.cell(row=row, column=1).value
        grid[week] = {
            'B1T1': ws.cell(row=row, column=2).value,
            'B1T2': ws.cell(row=row, column=3).value,
            'B2T1': ws.cell(row=row, column=4).value,
            'B2T2': ws.cell(row=row, column=5).value,
            'ByeT1': ws.cell(row=row, column=6).value,
            'ByeT2': ws.cell(row=row, column=7).value,
        }
    return grid


def load_requests(wb):
    ws = wb['Team Nights Off']
    reqs = []
    for row in range(5, 17):
        team = ws.cell(row=row, column=1).value
        night = ws.cell(row=row, column=3).value
        if team is None or str(team).strip() == '' or night is None or str(night).strip() == '':
            continue
        reqs.append((str(team).strip(), int(night)))
    return reqs


def find_role(row, team):
    for r, t in row.items():
        if t == team:
            return r
    return None


def opponent(row, team):
    role = find_role(row, team)
    if role is None:
        return None, None
    if role.startswith('Bye'):
        return None, role
    if role.startswith('B1'):
        other = 'B1T2' if role == 'B1T1' else 'B1T1'
    else:
        other = 'B2T2' if role == 'B2T1' else 'B2T1'
    return row[other], role


def solve(grid, requests):
    grid = {w: dict(r) for w, r in grid.items()}
    locked = {}   # night -> set of teams that MUST be on bye there
    log = []

    for X, N in requests:
        if N not in grid:
            log.append((X, N, 'ERROR', f"Night {N} is not a valid week (1-15)."))
            continue
        row_n = grid[N]
        role = find_role(row_n, X)
        if role and role.startswith('Bye'):
            log.append((X, N, 'OK', 'already on bye -- no change needed.'))
            locked.setdefault(N, set()).add(X)
            continue

        currently_locked = locked.get(N, set())

        if len(currently_locked) == 0:
            # Simple self-swap: whole-week trade with one of this team's
            # OTHER bye weeks, as long as that week isn't already locked
            # for someone else's request.
            candidates = [
                w for w in grid
                if w != N and w not in locked
                and (grid[w]['ByeT1'] == X or grid[w]['ByeT2'] == X)
            ]
            if not candidates:
                log.append((X, N, 'CONFLICT',
                            "No open week is available to swap with -- every "
                            "week where this team is on bye is already locked "
                            "by another request."))
                continue
            M = sorted(candidates)[0]
            grid[N], grid[M] = grid[M], grid[N]
            log.append((X, N, 'SWAPPED',
                        f"self-swap: traded places with week {M} (whole week moved)."))
            locked.setdefault(N, set()).add(X)

        else:
            # Night N already has a confirmed team on bye. Try to bring
            # this team in alongside them by borrowing the OTHER bye slot.
            p1, p2 = row_n['ByeT1'], row_n['ByeT2']
            borrow_candidates = [p for p in (p1, p2) if p not in currently_locked]
            if not borrow_candidates:
                log.append((X, N, 'CONFLICT',
                            "Both bye slots on this night are already locked "
                            "for other teams -- a 3rd team can't fit."))
                continue
            P = borrow_candidates[0]
            O_N, _ = opponent(row_n, X)
            if O_N is None:
                log.append((X, N, 'CONFLICT',
                            "Could not determine this team's opponent that night."))
                continue

            found_M = None
            for M in sorted(grid):
                if M == N or M in locked:
                    continue
                row_m = grid[M]
                if not (row_m['ByeT1'] == X or row_m['ByeT2'] == X):
                    continue
                p_opp, _ = opponent(row_m, P)
                if p_opp == O_N:
                    found_M = M
                    break

            if found_M is None:
                log.append((X, N, 'CONFLICT',
                            f"No valid compensating week found to borrow "
                            f"Team {P}'s bye slot without breaking someone "
                            f"else's total game count."))
                continue

            M = found_M
            for wk in (N, M):
                row = grid[wk]
                for r in row:
                    if row[r] == X:
                        row[r] = P
                    elif row[r] == P:
                        row[r] = X
            log.append((X, N, 'SWAPPED',
                        f"borrowed Team {P}'s bye slot -- traded roles with "
                        f"Team {P} on nights {N} and {M} to keep everyone's "
                        f"totals balanced."))
            locked.setdefault(N, set()).add(X)

    return grid, log


def verify(original_grid, solved_grid):
    """Sanity-check the solved grid against round-robin invariants."""
    from collections import defaultdict
    teams = sorted({t for row in original_grid.values() for t in row.values()})
    pair_count = defaultdict(int)
    bye_count = defaultdict(int)
    for row in solved_grid.values():
        m1 = tuple(sorted([row['B1T1'], row['B1T2']]))
        m2 = tuple(sorted([row['B2T1'], row['B2T2']]))
        pair_count[m1] += 1
        pair_count[m2] += 1
        bye_count[row['ByeT1']] += 1
        bye_count[row['ByeT2']] += 1
    problems = []
    for i in range(len(teams)):
        for j in range(i + 1, len(teams)):
            c = pair_count[(teams[i], teams[j])]
            if c != 2:
                problems.append(f"{teams[i]} vs {teams[j]} happens {c} times (should be 2)")
    for t in teams:
        if bye_count[t] != 5:
            problems.append(f"Team {t} has {bye_count[t]} byes (should be 5)")
    return problems


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 resolve_night_off.py <workbook.xlsx>")
        sys.exit(1)
    path = sys.argv[1]

    wb = openpyxl.load_workbook(path, data_only=False)
    if 'Base_Rotation' not in wb.sheetnames:
        print("ERROR: this workbook doesn't have a 'Base_Rotation' sheet yet.")
        print("Run setup_new_mechanism.py once first (ask Claude if you don't have it).")
        sys.exit(1)

    base_grid = load_base_rotation(wb)
    requests = load_requests(wb)

    if not requests:
        print("No Team Nights Off requests found -- nothing to solve.")
        return

    solved_grid, log = solve(base_grid, requests)

    problems = verify(base_grid, solved_grid)
    if problems:
        print("!! Refusing to save -- the solved schedule failed validation:")
        for p in problems:
            print("   -", p)
        sys.exit(1)

    # Write the solved grid into Reorg_Helper's resolved table (A2:G16),
    # keyed by week number (matches Schedule's direct VLOOKUP by night #).
    reorg = wb['Reorg_Helper']
    for week in range(1, 16):
        row_data = solved_grid[week]
        reorg.cell(row=week + 1, column=1, value=week)
        for i, role in enumerate(ROLES, start=2):
            reorg.cell(row=week + 1, column=i, value=row_data[role])

    wb.save(path)

    print(f"Solved {len(requests)} request(s):\n")
    for team, night, status, detail in log:
        print(f"  Team {team}, Night {night}: [{status}] {detail}")

    changed_weeks = sorted(
        w for w in base_grid if base_grid[w] != solved_grid[w]
    )
    print(f"\nWeeks changed from the original rotation: {changed_weeks or 'none'}")
    print("\nAll round-robin checks passed (every pair meets exactly twice, "
          "every team gets exactly 5 byes).")

    # Recalculate cached formula values so Schedule/Standings/etc. show up
    # correctly without needing Excel to reopen the file first.
    if shutil.which('soffice') or shutil.which('libreoffice'):
        try:
            subprocess.run(
                [sys.executable, '/mnt/skills/public/xlsx/scripts/recalc.py', path],
                check=True,
            )
            print("\nWorkbook recalculated.")
        except Exception as e:
            print(f"\n(Could not auto-recalculate: {e}. Just open the file in "
                  f"Excel once and it will refresh on its own.)")
    else:
        print("\n(LibreOffice not found on this machine -- just open the file "
              "in Excel once and it will recalculate automatically.)")


if __name__ == '__main__':
    main()
