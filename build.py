import zipfile
import xml.etree.ElementTree as ET
import json
import os
import re

XLSX_PATH = '比赛数据表.xlsx'
HEROES_DIR = 'heroes'
OUTPUT_PATH = 'data.js'

# BP correct time-sequence: column indices in game-time order
BP_ORDER = [
    # Phase 1: Ban 蓝→红→蓝→红
    (28, 'ban', '蓝'), (29, 'ban', '红'), (30, 'ban', '蓝'), (31, 'ban', '红'),
    # Phase 2: Pick 蓝→红红→蓝蓝→红
    (32, 'pick', '蓝'), (33, 'pick', '红'), (34, 'pick', '红'), (35, 'pick', '蓝'), (36, 'pick', '蓝'), (37, 'pick', '红'),
    # Phase 3: Ban 红→蓝→红→蓝→红→蓝
    (38, 'ban', '红'), (39, 'ban', '蓝'), (40, 'ban', '红'), (41, 'ban', '蓝'), (42, 'ban', '红'), (43, 'ban', '蓝'),
    # Phase 4: Pick 红→蓝蓝→红
    (44, 'pick', '红'), (45, 'pick', '蓝'), (46, 'pick', '蓝'), (47, 'pick', '红'),
]

MANUAL_HERO_MAP = {
    '云缨': '云樱',
    '铛': None,  # likely garbled, will fallback
}

def parse_xlsx(path):
    """Parse xlsx with stdlib zipfile + ElementTree, returns rows as list of list of strings."""
    with zipfile.ZipFile(path, 'r') as z:
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in tree.findall('.//ns:si', ns):
                    texts = []
                    for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                        if t.text:
                            texts.append(t.text)
                    shared_strings.append(''.join(texts))

        with z.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

            rows_data = []
            for row in tree.findall(f'.//{{{ns}}}row'):
                row_num = int(row.get('r'))
                if row_num == 1:
                    continue  # skip header

                cells = {}
                for c in row.findall(f'{{{ns}}}c'):
                    ref = c.get('r')
                    col_letter = re.match(r'([A-Z]+)', ref).group(1)
                    col_idx = col_to_idx(col_letter)
                    t = c.get('t', '')
                    v = c.find(f'{{{ns}}}v')
                    val = ''
                    if v is not None and v.text:
                        if t == 's':
                            idx = int(v.text)
                            if idx < len(shared_strings):
                                val = shared_strings[idx]
                        elif t == 'n':
                            try:
                                val = str(int(float(v.text)))
                            except ValueError:
                                val = v.text
                        else:
                            val = v.text
                    cells[col_idx] = val

                row_data = [cells.get(i, '') for i in range(48)]
                rows_data.append(row_data)

    return rows_data

def col_to_idx(col):
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def normalize_hero(name, image_set):
    if not name or not name.strip():
        return None, None

    name = name.strip()
    if name in MANUAL_HERO_MAP:
        mapped = MANUAL_HERO_MAP[name]
        if mapped is None:
            return name, None
        name = mapped

    if name + '.png' in image_set:
        return name, name + '.png'

    # half-width → full-width parentheses
    alt = name.replace('(', '（').replace(')', '）')
    if alt + '.png' in image_set:
        return alt, alt + '.png'

    # full-width → half-width parentheses
    alt2 = name.replace('（', '(').replace('）', ')')
    if alt2 + '.png' in image_set:
        return alt2, alt2 + '.png'

    return name, None  # will use default.png

def main():
    print('Reading xlsx...')
    rows = parse_xlsx(XLSX_PATH)
    print(f'  {len(rows)} data rows')

    print('Scanning heroes...')
    image_set = set(os.listdir(HEROES_DIR))
    print(f'  {len(image_set)} hero images')

    # Build hero image map
    hero_image_map = {}
    all_hero_names = set()
    for row in rows:
        for ci in [6, 8, 10, 12, 14, 19, 21, 23, 25, 27]:  # player hero columns
            name = row[ci].strip() if ci < len(row) else ''
            if name:
                all_hero_names.add(name)
        for ci in range(28, 48):  # BP hero columns
            name = row[ci].strip() if ci < len(row) else ''
            if name:
                all_hero_names.add(name)

    for h in sorted(all_hero_names):
        cleaned, img = normalize_hero(h, image_set)
        if cleaned:
            hero_image_map[cleaned] = img if img else 'default.png'
            # Also map original name if it differs
            if cleaned != h:
                hero_image_map[h] = img if img else 'default.png'
    hero_list = sorted(hero_image_map.keys())  # keep all variants for UI
    print(f'  {len(set(h for h in hero_list))} unique hero names in map')

    # Build teams, players
    team_player_map = {}
    player_team_map = {}
    all_players = set()

    for row in rows:
        for base in [2, 15]:  # team1 name at col 2, team2 name at col 15
            team = row[base].strip()
            if not team:
                continue
            if team not in team_player_map:
                team_player_map[team] = set()
            for offset in [3, 5, 7, 9, 11]:  # player ID columns relative to team base
                pid = row[base + offset].strip()
                if pid:
                    all_players.add(pid)
                    team_player_map[team].add(pid)
                    player_team_map[pid] = team

    teams = sorted([{'name': t, 'players': sorted(list(p))} for t, p in team_player_map.items()], key=lambda x: x['name'])
    players = sorted([{'fullId': p, 'team': player_team_map.get(p, ''), 'name': p.split('.')[-1] if '.' in p else p} for p in all_players], key=lambda x: x['fullId'])
    print(f'  {len(teams)} teams, {len(players)} players')

    print('Processing games...')
    games = []
    for row in rows:
        match_id = row[0]
        game_num = int(row[1]) if row[1] else 0

        t1_won = row[4].strip() == '胜'
        t2_won = row[17].strip() == '胜'

        team1 = {
            'name': row[2].strip(),
            'side': row[3].strip(),
            'won': t1_won,
            'players': [
                {'role': '对抗路', 'playerId': row[5].strip(), 'hero': row[6].strip()},
                {'role': '打野', 'playerId': row[7].strip(), 'hero': row[8].strip()},
                {'role': '中路', 'playerId': row[9].strip(), 'hero': row[10].strip()},
                {'role': '发育路', 'playerId': row[11].strip(), 'hero': row[12].strip()},
                {'role': '游走', 'playerId': row[13].strip(), 'hero': row[14].strip()},
            ]
        }
        team2 = {
            'name': row[15].strip(),
            'side': row[16].strip(),
            'won': t2_won,
            'players': [
                {'role': '对抗路', 'playerId': row[18].strip(), 'hero': row[19].strip()},
                {'role': '打野', 'playerId': row[20].strip(), 'hero': row[21].strip()},
                {'role': '中路', 'playerId': row[22].strip(), 'hero': row[23].strip()},
                {'role': '发育路', 'playerId': row[24].strip(), 'hero': row[25].strip()},
                {'role': '游走', 'playerId': row[26].strip(), 'hero': row[27].strip()},
            ]
        }

        # Build pick order from all pick columns in time sequence
        # Phase 2 pick column order: 蓝→红红→蓝蓝→红
        # Phase 4 pick column order: 红→蓝蓝→红
        pick_time_cols = [32, 33, 34, 35, 36, 37, 44, 45, 46, 47]  # time order
        pick_time_labels = [
            ('蓝',1),('红',1),('红',2),('蓝',2),('蓝',3),('红',3),
            ('红',4),('蓝',4),('蓝',5),('红',5)
        ]

        # Gather all 10 pick heroes with their order
        pick_heroes_in_order = []
        for global_idx, col_idx in enumerate(pick_time_cols):
            hero = row[col_idx].strip() if col_idx < len(row) else ''
            pick_heroes_in_order.append({
                'hero': hero,
                'globalOrder': global_idx + 1,
                'labelSide': pick_time_labels[global_idx][0],
                'labelNum': pick_time_labels[global_idx][1]
            })

        # Match each pick hero to a team and player
        def find_team_for_hero(hero_name):
            if not hero_name:
                return None, None
            for ti, t in enumerate([team1, team2]):
                for p in t['players']:
                    if p['hero'] == hero_name:
                        return ti, p['playerId']
            return None, None

        # Group picks by actual team
        t1_picks = []
        t2_picks = []
        for pe in pick_heroes_in_order:
            team_idx, pid = find_team_for_hero(pe['hero'])
            entry = {
                'hero': pe['hero'],
                'order': pe['globalOrder'],
                'playerId': pid or ''
            }
            if team_idx == 0:
                t1_picks.append(entry)
            elif team_idx == 1:
                t2_picks.append(entry)

        # Ban columns in time order: 蓝b1,红b1,蓝b2,红b2,红b3,蓝b3,红b4,蓝b4,红b5,蓝b5
        ban_time_cols = [28, 29, 30, 31, 38, 39, 40, 41, 42, 43]
        ban_time_sides = ['蓝','红','蓝','红','红','蓝','红','蓝','红','蓝']

        bp_seq = []
        ban_order = 0
        pick_order = len(ban_time_cols)  # bans come first in sequence numbering
        for i, col_idx in enumerate(ban_time_cols):
            hero = row[col_idx].strip() if col_idx < len(row) else ''
            side = ban_time_sides[i]
            ban_order += 1
            bp_seq.append({'order': ban_order, 'phase': 'ban', 'side': side, 'hero': hero})
        for i, col_idx in enumerate(pick_time_cols):
            hero = row[col_idx].strip() if col_idx < len(row) else ''
            side = pick_time_labels[i][0]
            pick_order += 1
            bp_seq.append({'order': pick_order, 'phase': 'pick', 'side': side, 'hero': hero})

        # Build bpByTeam based on actual team assignments for picks
        # Use ban column labeling for bans (trust the ban labels)
        bp_by_team = {}
        for ti, t in enumerate([team1, team2]):
            side_key = '蓝' if '蓝' in t['side'] else '红'
            bp_by_team[side_key] = {
                'bans': [{'hero': bp['hero'], 'order': bp['order']} for bp in bp_seq if bp['phase'] == 'ban' and bp['side'] == side_key],
                'picks': (t1_picks if ti == 0 else t2_picks)
            }

        date_str = match_id[:8] if len(match_id) >= 8 else match_id
        games.append({
            'matchId': match_id,
            'date': date_str,
            'gameNum': game_num,
            'team1': team1,
            'team2': team2,
            'bpSequence': bp_seq,
            'bpByTeam': bp_by_team
        })

    print(f'  {len(games)} games processed')

    # Compute meta
    match_ids = sorted(set(g['matchId'] for g in games))
    dates = sorted(set(g['date'] for g in games))
    print('Writing data.js...')
    data = {
        'meta': {
            'totalGames': len(games),
            'totalMatches': len(match_ids),
            'totalTeams': len(teams),
            'totalPlayers': len(players),
            'totalHeroes': len(hero_list),
            'dateMin': dates[0] if dates else '',
            'dateMax': dates[-1] if dates else '',
            'matchIdMin': match_ids[0] if match_ids else '',
            'matchIdMax': match_ids[-1] if match_ids else '',
        },
        'heroImageMap': hero_image_map,
        'teams': teams,
        'players': players,
        'heroes': hero_list,
        'games': games,
    }

    json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write('window.GAME_DATA=')
        f.write(json_str)
        f.write(';')

    size_mb = os.path.getsize(OUTPUT_PATH) / 1024 / 1024
    print(f'Done! Output: {OUTPUT_PATH} ({size_mb:.2f} MB)')

if __name__ == '__main__':
    main()
