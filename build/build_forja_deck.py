# -*- coding: utf-8 -*-
"""Сборка доклада «Forja: от вайб-кодинга к продукту» в стиле forja (stylekit AgentKit).

Запуск:      python -X utf8 build_forja_deck.py
Результат:   2026-11-01_forja-выпускной_доклад.pptx рядом со скриптом (файл перезаписывается).
Проверка:    python -X utf8 <agentkit>/src/agentkit/skills/pptx/scripts/pp.py check <pptx>
             python -X utf8 <agentkit>/src/agentkit/skills/pptx/scripts/pp.py png <pptx> --out <папка>

Правила сборки: 16:9, координаты в дюймах от st.content, цвета и шрифты - ролями токенов стиля,
картинок нет (только фигуры и текст). Ниже у каждого слайда комментарий: какая схема и почему такие
координаты. Часть схем собрана здесь, а не в stylekit: у библиотечных компонентов кегли 9-11 пт,
зазоры 0,1 дюйма между стрелкой и блоком и подпись плашкой поверх линии - по приёмке доклада так
нельзя (не мельче 12 пт, стрелка упирается в блок, подпись стоит рядом с линией).
"""
import math
import sys
from pathlib import Path

# stylekit плагина: корень пакета ставится в sys.path, иначе подхватится чужая установка стилей
AGENTKIT = Path(r'C:\Users\khser\AIProjects\Products\AgentKit\src\agentkit')
sys.path.insert(0, str(AGENTKIT))

from pptx.enum.dml import MSO_LINE_DASH_STYLE  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

from stylekit import icons, load_style, objects  # noqa: E402

OUT = Path(__file__).with_name('2026-11-01_forja-выпускной_доклад.pptx')
C, M = PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE

st = load_style('forja')
prs = st.deck()
CW = st.w - 2 * st.m          # рабочая ширина, 12,55 дюйма
TOP = 2.18                    # низ шапки слайда с подзаголовком (кикер + заголовок + линия + sub)
FOOT = 6.756                  # верх подвала (номер страницы, сноска)


def conclusion(s, head, body, y=5.60, h=0.95):
    """Плашка вывода во всю ширину внизу слайда: заголовок и одна-две строки текста."""
    st.callout(s, st.m, y, CW, h, head, body)


def conclusion_marked(s, head, body, y, h):
    """Плашка вывода с меткой «!»: знак, заголовок и текст внутри полосы, слева.
    Зачем не st.alert: тот выносит заголовок над полосой и вешает знак справа - на остальных
    слайдах знак и заголовок стоят в полосе слева."""
    st.rect(s, st.m, y, CW, h, fill='callout')
    st.rect(s, st.m, y, 0.05, h, fill='accent')
    st.mark(s, st.m + 0.22, y + 0.14)
    st.text(s, st.m + 0.58, y + 0.13, CW - 0.8, 0.30, head, size='callout_head', color='title', bold=True)
    st.text(s, st.m + 0.58, y + 0.50, CW - 0.8, h - 0.60, body, size='body', color='text', spacing=1.1,
            bold_color='title')


def notes(s, txt):
    s.notes_slide.notes_text_frame.text = txt


def drop_shadows(s):
    """Гасит тень у всех фигур слайда. У коннекторов тень наследуется от темы PowerPoint и на тонкой
    линии читается серым размытым дублем; st.rect гасит её сам, st.line и st.arrow - нет."""
    for sh in s.shapes:
        try:
            sh.shadow.inherit = False
        except (AttributeError, ValueError):
            pass


def cards(st_, s, x, y, w, h, items, cols=3, gap=0.28, badge=0.62):
    """Карточки с пиктограммой, высота задаётся под текст, содержимое центровано по вертикали.
    Зачем не icons.icon_cards: там карточка тянется на всю область, и под двухстрочным текстом
    остаётся пустая треть (0,7 дюйма на карточке 2,5)."""
    cw = (w - gap * (cols - 1)) / cols
    block = badge + 0.12 + 0.36 + 0.10 + 0.42       # значок, заголовок, две строки текста
    top = max(0.16, (h - block) / 2)
    for i, it in enumerate(items):
        name, head = it[0], it[1]
        body = it[2] if len(it) > 2 else None
        cx = x + (i % cols) * (cw + gap)
        cy = y + (i // cols) * (h + gap)
        st.rect(s, cx, cy, cw, h, fill='surface', line='line')
        icons.icon_badge(st, s, cx + 0.22, cy + top, badge, name)
        st.text(s, cx + 0.22, cy + top + badge + 0.12, cw - 0.44, 0.36, head, size='card_head',
                color='title', bold=True)
        if body:
            body_y = cy + top + badge + 0.58
            st.text(s, cx + 0.22, body_y, cw - 0.44, cy + h - body_y - 0.14, body, size='small',
                    color='text', spacing=1.1)


def hex_chain(st_, s, x, y, w, h, items):
    """Соты-этапы в ряд: светлый контейнер, синяя сота с названием, пояснение под ней.
    Зачем не st.hex_flow: тот рисует вводную стрелку, которая начинается в пустоте и до первой
    соты не доходит, и ставит пояснение 11 пт."""
    n = len(items)
    cw = w / n
    ow, hw, hh = 2.55, 1.5, 1.3
    cy = y + 0.26 + hh / 2                          # центр синей соты
    for i, (name, txt) in enumerate(items):
        cx = x + i * cw + cw / 2
        st.vhex(s, cx, y + h / 2, ow, h, fill='surface', line='line', lw=1.0)
        st.vhex(s, cx, cy, hw, hh, fill='primary')
        st.text(s, cx - hw / 2, cy - hh / 2, hw, hh, name, size=16, color='white', bold=True,
                caps=True, align=C, anchor=M)
        ty = cy + hh / 2 + 0.12
        st.text(s, cx - 1.02, ty, 2.04, y + h - ty - 0.14, txt, size=13, color='text', align=C,
                spacing=1.05)
        if i < n - 1:                               # стрелка от края контейнера до края соседнего
            st.arrow(s, cx + ow / 2, cy, cx + cw - ow / 2, cy, color='primary', lw=3)


def slide_sub(st_, s, y, txt):
    """Подзаголовок шапки в две строки. Зачем своё: st.slide отводит под sub 0,26 дюйма,
    длинная фраза из этой рамки вылезает, и pp.py check считает это ошибкой."""
    st_.text(s, st_.m, y, st_.w - 2 * st_.m, 0.56, txt, size='subtitle', color='text_soft')


def circle_cycle(st_, s, cx, cy, rx, ry, names, node_w=2.05, node_h=0.62, start=-90):
    """Цикл кольцом: этапы по эллипсу, дуги-стрелки по часовой стрелке, последний замыкается на первый.
    Зачем своё: st.hex_flow рисует цикл цепочкой с двумя разворотами по краям, а не кольцом.
    Дуги начинаются от края блока и упираются в край следующего: концы спрятаны под плашками.
    start - угол первого этапа: -90 ставит его наверх, -60 разводит этапы по шести точкам эллипса, и
    у всех шести пар остаётся длинная дуга (при этапе наверху боковые пары стоят вплотную)."""
    n = len(names)
    step = 360.0 / n

    def pt(t):                                     # точка эллипса, 0 - справа, 90 - снизу
        a = math.radians(t)
        return cx + rx * math.cos(a), cy + ry * math.sin(a)

    def box(i):
        px, py = pt(start + i * step)              # от первого этапа дальше по часовой
        return px - node_w / 2, py - node_h / 2

    def inside(t, i, grow=0.0):
        px, py = pt(t)
        bx, by = box(i)
        return bx - grow <= px <= bx + node_w + grow and by - grow <= py <= by + node_h + grow

    for i in range(n):
        a0, a1 = start + i * step, start + (i + 1) * step
        t0 = a0
        while t0 < a1 - 2 and inside(t0, i, 0.06):           # вышли из своей плашки
            t0 += 0.5
        t1 = a1
        while t1 > t0 + 2 and inside(t1, (i + 1) % n):       # от центра следующей плашки назад до её края
            t1 -= 0.5
        t1 -= 0.5                                            # кончик острия - у самого края, не под плашкой
        arc = [pt(t0 + (t1 - t0) * j / 18) for j in range(19)]
        fb = s.shapes.build_freeform(Inches(arc[0][0]), Inches(arc[0][1]), scale=1.0)
        fb.add_line_segments([(Inches(px), Inches(py)) for px, py in arc[1:]], close=False)
        shp = fb.convert_to_shape()
        shp.fill.background()
        shp.line.color.rgb = st_.c('accent')
        shp.line.width = Pt(2.25)
        shp.shadow.inherit = False
        (bx, by), (tx, ty) = pt(t1 - 4), arc[-1]    # остриё по касательной, кончик - на краю плашки
        dx, dy = tx - bx, ty - by
        ln = math.hypot(dx, dy) or 1
        ux, uy = dx / ln, dy / ln
        st_.poly(s, [(tx, ty), (tx - ux * 0.16 - uy * 0.075, ty - uy * 0.16 + ux * 0.075),
                     (tx - ux * 0.16 + uy * 0.075, ty - uy * 0.16 - ux * 0.075)], fill='accent')
    for i, name in enumerate(names):                       # плашки поверх концов дуг
        bx, by = box(i)
        st_.rect(s, bx, by, node_w, node_h, fill='primary', radius=0.10)
        st_.text(s, bx + 0.1, by, node_w - 0.2, node_h, name, size='scheme', color='white', bold=True,
                 align=C, anchor=M, spacing=0.95)


def safeguards(st_, s, x, y, w, h, items, gap=0.22):
    """Контур сохранности: три плашки в ряд, у каждой название и пояснение."""
    cw = (w - gap * (len(items) - 1)) / len(items)
    for i, (head, txt) in enumerate(items):
        cx = x + i * (cw + gap)
        st_.rect(s, cx, y, cw, h, fill='surface', line='line')
        st_.rect(s, cx, y, 0.05, h, fill='accent')
        st_.text(s, cx + 0.22, y + 0.06, cw - 0.4, 0.26, head, size='card_head', color='title', bold=True)
        st_.text(s, cx + 0.22, y + 0.32, cw - 0.4, h - 0.40, txt, size='small', color='text_soft')


def step_flow(st_, s, x, y, w, h, steps, note=None):
    """Шаги процесса: блоки во всю высоту панели, роль белой плашкой внизу блока, стрелки вплотную.
    Зачем не objects_process.process_scheme: там роль 10 пт, стрелка 0,3 дюйма цвета «ось» с тенью,
    а между названием и плашкой роли остаётся 0,4 дюйма пустоты."""
    st.rect(s, x, y, w, h, fill='surface', line='line')
    n = len(steps)
    aw = 0.55                                       # длина стрелки = зазор между блоками
    bw = (w - 0.36 - aw * (n - 1)) / n
    bh = h - 1.35
    by = y + 0.35
    bx = x + 0.18
    for i, (name, role) in enumerate(steps):
        st.rect(s, bx, by, bw, bh, fill='primary', radius=0.05)
        st.text(s, bx + 0.1, by + 0.08, bw - 0.2, bh - 0.68, name, size=14, color='white',
                bold=True, align=C, anchor=M, spacing=1.0)
        pw = min(bw - 0.24, 1.2)
        st.rect(s, bx + (bw - pw) / 2, by + bh - 0.42, pw, 0.32, fill='white', radius=0.16)
        st.text(s, bx + (bw - pw) / 2, by + bh - 0.42, pw, 0.32, role, size='small',
                color='primary', align=C, anchor=M)
        if i < n - 1:
            st.arrow(s, bx + bw, by + bh / 2, bx + bw + aw, by + bh / 2, color='primary', lw=2.5)
        bx += bw + aw
    if note:
        st.mark(s, x + 0.3, y + h - 0.72)
        st.text(s, x + 0.7, y + h - 0.78, w - 1.2, 0.4, note[1], size='card_head',
                color='accent_text', bold=True, anchor=M)


def before_after(st_, s, x, y, w, h, rows, aspect_w=2.4):
    """Было → стало: строка - аспект, серая плашка, стрелка, голубая плашка.
    Зачем своё: st.before_after набирает плашки 11 пт и оставляет 0,1 дюйма до блоков."""
    cw = (w - aspect_w - 1.0) / 2
    bx = x + aspect_w + 0.1
    ax = bx + cw + 0.7
    st.text(s, bx, y, cw, 0.4, 'Было', size=14, color='muted', bold=True, caps=True)
    st.text(s, ax, y, cw, 0.4, 'Стало', size=14, color='primary', bold=True, caps=True)
    top = y + 0.55
    rh = (h - 0.55) / len(rows)
    for i, (asp, was, now) in enumerate(rows):
        ry = top + i * rh
        bh = rh - 0.16
        st.text(s, x, ry, aspect_w, bh, asp, size=13, color='primary', bold=True, caps=True, anchor=M)
        st.rect(s, bx, ry, cw, bh, fill='surface')
        st.text(s, bx + 0.18, ry, cw - 0.36, bh, was, size=14, color='text_soft', anchor=M,
                spacing=1.0)
        st.arrow(s, bx + cw, ry + bh / 2, ax, ry + bh / 2, color='primary', lw=3)
        st.rect(s, ax, ry, cw, bh, fill='light')
        st.text(s, ax + 0.18, ry, cw - 0.36, bh, now, size=14, color='primary', anchor=M,
                spacing=1.0)


def hub(st_, s, cx, cy, rx, ry, center, nodes, pw=1.9, ph=0.5):
    """«Хаб и спицы»: центр и шесть узлов по эллипсу, пунктирные спицы.
    Зачем своё: в objects_practice.hub_spoke подписи узлов набраны 10 пт."""
    pts = []
    for name, deg in nodes:
        a = math.radians(deg)
        pts.append((name, cx + rx * math.cos(a), cy + ry * math.sin(a)))
    for name, nx, ny in pts:                        # спицы под плашками: концов не видно, разрывов нет
        dx, dy = nx - cx, ny - cy
        ln = math.hypot(dx, dy) or 1
        ux, uy = dx / ln, dy / ln
        k = 1 / max(abs(ux) / (pw / 2), abs(uy) / (ph / 2))
        ln_sh = st.line(s, cx, cy, nx - ux * k, ny - uy * k, color='axis', lw=1.1)
        ln_sh.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    for name, nx, ny in pts:
        st.rect(s, nx - pw / 2, ny - ph / 2, pw, ph, fill='surface', line='line', radius=0.22)
        st.text(s, nx - pw / 2, ny - ph / 2, pw, ph, name, size=13, color='text', align=C, anchor=M)
    st.rect(s, cx - 0.8, cy - 0.55, 1.6, 1.1, fill='primary', shape=MSO_SHAPE.OVAL)
    st.text(s, cx - 0.8, cy - 0.55, 1.6, 1.1, center, size='card_head', color='white', bold=True,
            align=C, anchor=M)


def interaction_scheme(st_, s, x, y, w, h, nodes, flows, bw=2.3, bh=0.95):
    """Кто с кем чем обменивается: блоки по местам, стрелки от границы блока до границы блока.
    Зачем своё: st.interaction подписывает поток плашкой цвета фона прямо на линии и рвёт её на два
    обрезка, у наклонных потоков обрывки не совпадают по прямой, а до блоков остаётся 0,08 дюйма."""
    pos = [(x + bw / 2 + n['x'] * (w - bw), y + bh / 2 + n['y'] * (h - bh)) for n in nodes]
    pairs = {(f[0], f[1]) for f in flows}
    for fl in flows:
        a, b, lab = fl[0], fl[1], fl[2]
        (ax, ay), (bx, by) = pos[a], pos[b]
        dx, dy = bx - ax, by - ay
        ln = math.hypot(dx, dy) or 1
        ux, uy = dx / ln, dy / ln
        ox = oy = 0.0
        if (b, a) in pairs:                         # встречный поток отводим вбок, иначе ляжет на линию
            ox, oy = -uy * 0.13, ux * 0.13
        k = 1 / max(abs(ux) / (bw / 2), abs(uy) / (bh / 2))
        x1, y1 = ax + ox + ux * k, ay + oy + uy * k
        x2, y2 = bx + ox - ux * k, by + oy - uy * k
        st.arrow(s, x1, y1, x2, y2, color='primary', lw=1.75)
        if lab:                                     # подпись сбоку от линии: она остаётся целой
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            if oy:                                  # встречная пара: подпись уходит наружу от пары
                lx, ly = mx - 0.7, my + (0.22 if oy > 0 else -0.22) - 0.13
            else:
                lx, ly = mx + uy * 0.22 - 0.7, my - ux * 0.22 - 0.13
            st.text(s, lx, ly, 1.4, 0.26, lab, size='small', color='primary', align=C, anchor=M)
    for n, (cx, cy) in zip(nodes, pos):
        bx, by = cx - bw / 2, cy - bh / 2
        kind = n.get('kind', 'party')
        if kind == 'center':
            st.rect(s, bx, by, bw, bh, fill='primary')
            tc = 'white'
        elif kind == 'system':
            shp = st.rect(s, bx, by, bw, bh, fill='light', line='primary', lw=1.0)
            shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            tc = 'primary'
        else:
            st.rect(s, bx, by, bw, bh, fill='card')
            st.rect(s, bx, by, bw, 0.07, fill='sky')
            tc = 'primary'
        st.text(s, bx + 0.1, by + 0.10, bw - 0.2, 0.38, n['name'], size=14, color=tc, bold=True,
                caps=True, align=C, font='numbers')
        if n.get('note'):
            st.text(s, bx + 0.1, by + 0.46, bw - 0.2, 0.42, n['note'], size='small', color=tc,
                    align=C, spacing=0.95)


def done_plans(st_, s, x, y, w, h, done, plans, gap=0.55):
    """Две колонки: слева «сделано» с галочками (название и пояснение), справа шаги со сроком.
    Зачем своё: st.next_steps даёт одну колонку на всю ширину и стрелки с тенью."""
    lw = (w - gap) / 2
    rx = x + lw + gap
    st_.text(s, x, y, lw, 0.32, 'Сделано за курс', size='small', color='primary', bold=True,
             caps=True, tracking=0.5)
    st_.text(s, rx, y, lw, 0.32, 'Дальше', size='small', color='primary', bold=True, caps=True,
             tracking=0.5)
    rh = (h - 0.46 - 0.10 * (len(done) - 1)) / len(done)
    for i, (head, body) in enumerate(done):
        yy = y + 0.46 + i * (rh + 0.10)
        st_.rect(s, x, yy, lw, rh, fill='surface', line='line')
        icons.icon_badge(st_, s, x + 0.2, yy + (rh - 0.34) / 2, 0.34, 'check', fill='good', size=10)
        st_.text(s, x + 0.74, yy + 0.08, lw - 0.96, 0.26, head, size=14, color='title', bold=True)
        st_.text(s, x + 0.74, yy + 0.37, lw - 0.96, rh - 0.45, body, size='small', color='text',
                 spacing=1.05)
    step = (h - 0.46) / len(plans)
    for i, (when, what) in enumerate(plans):
        yy = y + 0.46 + i * step
        st_.text(s, rx, yy + 0.02, 1.35, 0.34, when, size=13, color='primary', bold=True)
        st_.text(s, rx + 1.4, yy, lw - 1.4, step - 0.06, what, size=13, color='text', spacing=1.05)
        if i < len(plans) - 1:
            st_.line(s, rx, yy + step - 0.04, rx + lw, yy + step - 0.04, color='line', lw=0.75)


def funnel_bands(s, x, y, w, h, stages, result=None):
    """Воронка этапов: сужающиеся полосы, кольца на границах, подписи сверху и снизу конуса.

    Почему не st.funnel: там подпись этапа стоит в рамке высотой 0,3 дюйма, а двухстрочные названия
    («Подключил модель», «Помощники и навыки») из неё вылезают на 7 пт - pp.py check считает это ошибкой.
    Здесь та же геометрия (полосы с градиентом, кольца, плашка результата), но у подписи своя высота,
    а разделители колонок начинаются под подписями, а не поверх них.
    """
    n = len(stages)
    wf = w - (2.4 if result else 0)
    sw = wf / n
    head_h, big, small = 0.62, 1.84, 0.62
    cy = y + head_h + 0.10 + big / 2                  # центр конуса
    desc_y = cy + big / 2 + 0.14
    for i, (head, txt) in enumerate(stages):
        xl, xr = x + i * sw, x + (i + 1) * sw
        hl = big - (big - small) * i / n
        hr = big - (big - small) * (i + 1) / n
        band = st.poly(s, [(xl, cy - hl / 2), (xr, cy - hr / 2), (xr, cy + hr / 2), (xl, cy + hl / 2)],
                        fill='light')
        st.gradient(band, 'light', 'white', angle=90)
        band.line.color.rgb, band.line.width = st.c('line'), Pt(0.75)
        st.rect(s, xl - 0.11, cy - hl / 2, 0.22, hl, fill='primary', shape=MSO_SHAPE.OVAL)
        if i:
            st.line(s, xl, y + head_h + 0.06, xl, y + h, color='axis', lw=0.5)
        st.text(s, xl + 0.05, y, sw - 0.1, head_h, head, size='small', color='primary', bold=True,
                caps=True, align=C, spacing=0.95)
        st.text(s, xl + 0.05, desc_y, sw - 0.1, y + h - desc_y, txt, size='small', color='text_soft',
                align=C, spacing=0.95)
    st.rect(s, x + wf - 0.11, cy - small / 2, 0.22, small, fill='sky', shape=MSO_SHAPE.OVAL)
    if result:
        rx = x + wf + 0.42
        st.arrow(s, x + wf + 0.02, cy, rx - 0.08, cy, color='accent', lw=2.5)
        st.rect(s, rx, y + 0.1, x + w - rx, h - 0.1, fill='sky')
        st.text(s, rx + 0.15, y + 0.35, x + w - rx - 0.3, 0.6, result[0], size='card_head', color='white',
                bold=True, caps=True, font='numbers')
        st.text(s, rx + 0.18, y + 1.20, x + w - rx - 0.36, h - 1.40, result[1], size=13,
                color='white', spacing=1.05)


# ---------------------------------------------------------------- 1. титул
s = st.cover(prs, 'Forja: от вайб-кодинга к продукту',
             subtitle='Рабочее пространство для специалистов: свои знания, готовые помощники, '
                      'лучшие модели без настройки',
             date='01.11.2026', kicker='ZeroCoder · онлайн-выпускной',
             author='Хачатрян Сергей Арамович', position='Руководитель проектов')
notes(s, 'Приветствие. Одна фраза о пути: за два месяца — от первого скрипта до продукта '
         'с инвестиционной идеей. Дальше — шесть блоков.')
drop_shadows(s)

# ---------------------------------------------------------------- 2. проблема
# cards: три карточки в ряд; высота 2,55 под две строки текста, содержимое центровано.
# slide_sub: подзаголовок в две строки, поэтому рисуется своим блоком 0,56 дюйма, а не полем sub.
s = st.slide(prs, 'Проблема', 'ИИ был собеседником, стал исполнителем')
slide_sub(st, s, 1.78, 'Раньше — чат: спросил, получил ответ, работу делал сам. Сегодня — '
                       'инструмент, который знает мои документы, правила и проекты и доводит задачу '
                       'до конца.')
cards(st, s, st.m, 2.40, CW, 2.30, [
    ('doc', 'Было: чат',
     'Спросил — ответил. Контекст терялся, каждый раз с нуля.'),
    ('search', 'Стало: среда',
     'Библиотека знаний: документы, правила, история решений.'),
    ('lock', 'Разница',
     'Дело не в модели, а в том, к каким знаниям она подключена.'),
], cols=3, gap=0.28)
conclusion(s, 'Вывод', 'Работа не просто ускорилась — она перешла от меня к агенту.',
           y=5.15, h=1.05)
notes(s, '15 лет трансформационных проектов, команда из 5 аналитиков, материалы идут руководству '
         'корпорации. Каждый документ — ручная сборка.')
drop_shadows(s)

# ---------------------------------------------------------------- 3. две недели
# hex_chain: четыре соты в ряд; контейнер 2,55 x 2,45, синяя сота 1,5 x 1,3, подпись 13 pt.
s = st.slide(prs, 'Блок 1 · Знакомство', 'Две недели вайб-кодинга',
             sub='От скрипта до первого агента — каждый шаг в реальной работе')
hex_chain(st, s, st.m, 2.52, CW, 2.45, [
    ('Скрипт', 'Выгрузка Excel без ручной сборки'),
    ('Сайт', 'Лендинг услуг за вечер'),
    ('Бот', 'Отвечает в Telegram вместо меня'),
    ('Агент', 'Задача на входе, файл на выходе'),
])
conclusion(s, 'Что понял', 'Писать код перестало быть узким местом. Узким местом стали данные.',
           y=5.55)
st.footnote(s, 'ZeroCoder: 12 уроков, 8 домашних проектов')
notes(s, 'Показать, что первые проекты были учебные, но делались на своих задачах: лендинг '
         'консалтинга, разбор выгрузки, бот-тренажёр.')
drop_shadows(s)

# ---------------------------------------------------------------- 4. слой данных
# circle_cycle: шесть этапов кольцом по эллипсу (4,9 x 1,05) с угла -60: этапы стоят по шести точкам,
# дуги-стрелки по часовой, замыкание последнего на первый идёт по верхней дуге; подпись этапа 13 pt.
# safeguards: три плашки контура сохранности под кольцом, 0,56 дюйма на название и пояснение.
s = st.slide(prs, 'Блок 2 · Слой данных', 'Знание возвращается в библиотеку',
             sub='Материал проходит круг: разбор, обезличивание, индекс, работа, новое знание — и снова в библиотеку')
circle_cycle(st, s, st.m + CW / 2, 3.45, 4.90, 1.05, [
    'Материал', 'ingest: разбор', 'Обезличивание', 'Индекс', 'Работа агента', 'Новое знание',
], start=-60)
st.text(s, st.m + CW / 2 - 2.4, 3.27, 4.8, 0.36, 'Библиотека знаний', size='card_head',
        color='primary', bold=True, align=C)
safeguards(st, s, st.m, 4.88, CW, 0.56, [
    ('Хуки', 'проверяют каждое действие'),
    ('Страж', 'не пускает закрытое наружу'),
    ('Бэкап', 'история версий'),
])
conclusion(s, 'Правило', 'Оригинал лежит в raw и не меняется. Работа идёт с копией.', y=5.58)
notes(s, 'Ключевая мысль блока: слой данных — это то, что делает агента полезным. Модель '
         'подставляется любая.')
drop_shadows(s)

# ---------------------------------------------------------------- 5. модель - деталь
# gpcps_business.service_model: четыре слоя сверху вниз, между слоями пара стрелок с меткой.
# Столбец подписей 2,15", правая колонка 1,35" оставлена под отдельную связь «Агенты → Модели»:
# в схеме связь рисуется только между соседними слоями, а «вызов» идёт через слой данных.
# Элементы слоя в схеме набраны 11 pt - после сборки поднимаем кегль до 12.
s = st.slide(prs, 'Блок 2 · Устройство среды', 'Модель — сменная деталь',
             sub='Среда держится на данных, а не на вендоре')
SM_W = CW - 1.35
_lo = len(s.shapes)
st.service_model(s, [
    ('Интерфейс', ['Forja', 'чат с агентом', 'проекты', 'файлы']),
    ('Агенты', ['7 субагентов', '25 скиллов', '10 хуков']),
    ('Данные', ['raw', 'wiki', 'индекс', 'библиотека']),
    ('Модели и инструменты', ['свой шлюз', '4 MCP', 'любой провайдер']),
], links=['задача', 'чтение и запись', None], x=st.m, y=TOP, w=SM_W, h=3.30, label_w=2.15)
for _sh in list(s.shapes)[_lo:]:
    if not _sh.has_text_frame:
        continue
    for _par in _sh.text_frame.paragraphs:
        for _run in _par.runs:
            if _run.font.size and _run.font.size.pt < 12:
                _run.font.size = Pt(12)
_lh = (3.30 - 0.42 * 3) / 4                       # высота слоя в service_model
_y1 = TOP + _lh / 2 + (_lh + 0.42)               # центр слоя «Агенты»
_y3 = TOP + _lh / 2 + 3 * (_lh + 0.42)           # центр слоя «Модели и инструменты»
_xe, _xl = st.m + SM_W, st.m + SM_W + 0.35       # край слоёв и отступ обходной связи
st.line(s, _xe, _y1, _xl, _y1, color='primary', lw=1.5)
st.line(s, _xl, _y1, _xl, _y3, color='primary', lw=1.5)
st.arrow(s, _xl, _y3, _xe + 0.02, _y3, color='primary', lw=1.5)
# метка «вызов» на обходной линии: 12 pt, шестигранник по центру линии
_ty = (_y1 + _y3) / 2 - 0.15
st.rect(s, _xl - 0.45, _ty, 0.9, 0.30, fill='deep', shape=MSO_SHAPE.HEXAGON)
st.text(s, _xl - 0.45, _ty, 0.9, 0.30, 'вызов', size='small', color='white', align=C, anchor=M)
conclusion(s, 'Проверка на практике',
           'Провайдер закрыл доступ — среда переехала на свой шлюз за день. Проекты не пересобирались.',
           y=5.55, h=1.0)
st.footnote(s, 'Источник: разведка среды ~/.claude, 01.11.2026')
notes(s, 'Здесь главный аргумент: агент и модель могут быть любыми. Ушёл Anthropic — работали '
         'дальше на других моделях через локальный шлюз.')
drop_shadows(s)

# ---------------------------------------------------------------- 6. рутина
# before_after: строка = аспект, «было» и «стало» плашками 14 pt, стрелка упирается в блоки.
s = st.slide(prs, 'Блок 3 · Автоматизация', 'Рутину забрали агенты',
             sub='Сначала дорого и бесполезно. Потом — роли, файлы, проверки.')
before_after(st, s, st.m, 2.20, CW, 3.15, [
    ('Документ', 'Часы вручную', 'Задание агенту'),
    ('Расчёт', 'Excel руками', 'Книга расчётов по шаблону'),
    ('Проверка', 'На глаз', 'Хуки и агент качества'),
    ('Повтор', 'Заново каждый раз', 'Воспроизводимо'),
], aspect_w=2.4)
conclusion(s, 'Эффект', 'OmniScope и AgentKit: доклад, расчёт, документ, схема', y=5.55)
notes(s, 'Честно сказать: первые попытки были дорогими и слабыми. Перелом — когда работа '
         'разложилась на роли и описалась файлами.')
drop_shadows(s)

# ---------------------------------------------------------------- 7. своя CRM
# hub: центр и шесть узлов по эллипсу (rx 4,35, ry 1,45), подписи узлов 13 pt.
s = st.slide(prs, 'Блок 4 · Приложение', 'Своя CRM: CoreDesk → Forja',
             sub='Приложение живёт там, где им пользуются каждый день')
hub(st, s, st.m + CW / 2, 3.82, 4.35, 1.45, 'Forja', [
    ('Задачи', 0), ('Клиенты', 35), ('Файлы', 145),
    ('Агент', 180), ('Календарь', 215), ('Доска', 325),
])
conclusion(s, 'Как получилось', 'Не одним релизом: итерации и параллельное развитие среды.',
           y=5.65, h=0.90)
notes(s, 'Креативность и старательность: продукт рождался из ежедневного неудобства, а не из '
         'плана на год.')
drop_shadows(s)

# ---------------------------------------------------------------- 8. что работает
# icons.icon_cards: шесть карточек 3 x 2; у каждой строки описание, бейдж 0,46".
s = st.slide(prs, 'Блок 5 · Продукты', 'Что уже работает',
             sub='Не прототипы — то, чем пользуются.')
icons.icon_cards(st, s, st.m, TOP, CW, 3.42, [
    ('globe', 'Сайты', 'лендинги услуг за вечер'),
    ('send', 'Автопостинг в Telegram', 'анонсы в канал по расписанию'),
    ('calendar', 'TayaBot', 'запись к мастеру'),
    ('checklist', 'PMPTrainer', 'тренажёр к экзамену'),
    ('grid', 'Teamboard', 'доска проектов'),
    ('network', 'Свой сервер', 'VPS'),
], cols=3, gap=0.28, badge=0.46)
conclusion(s, 'Вывод', 'Каждый продукт — кусочек пазла, а не хобби в стороне от работы.',
           y=5.75, h=0.90)
notes(s, 'Перечислить быстро, один пример показать живьём: TayaBot или автопостинг.')
drop_shadows(s)

# ---------------------------------------------------------------- 9. демонстрация
# objects.numbered_rows: номер в шевроне, текст строкой; имя продукта выделено жирным разметкой стиля.
s = st.slide(prs, 'Живой показ', 'Демонстрация', sub='Четыре точки, каждая — рабочий результат.')
objects.numbered_rows(st, s, st.m, 2.25, CW, [
    '**Forja** — проект, задача, агент: результат в папке проекта',
    '**Скилл** — одно задание: доклад и книга расчётов',
    '**TayaBot** — запись к мастеру за полминуты',
    '**Репозиторий** — код, README, скриншоты, скилл',
], row_h=0.72, gap=0.12)
conclusion(s, 'Страховка', 'Если не запустится — готовые результаты лежат в репозитории.',
           y=5.70, h=0.90)
notes(s, 'Держать заготовленные результаты открытыми заранее: скриншоты в README, папка с докладом '
         'и книгой.')
drop_shadows(s)

# ---------------------------------------------------------------- 10. пазл
# Решётка 2 x 2: номер и стрелка сверху, заголовок и пояснение сразу под ними (без пустого промежутка).
s = st.slide(prs, 'Блок 6 · Итог', 'Пазл сложился: Forja AI Studio',
             sub='Среда, студия, продукт и доска — четыре части одной идеи')
_tiles = [('Forja', 'Имя среды: задача превращается в результат'),
          ('Forja AI Studio', 'Студия: кто делает и по какой методе'),
          ('Forja Anvil', 'Продукт: рабочее место человека и агентов'),
          ('Forja Board', 'Доска, где задача уходит агенту')]
_tw, _th, _tg = (CW - 0.3) / 2, 1.55, 0.18
for _i, (_head, _note) in enumerate(_tiles):
    _cx, _cy = st.m + (_i % 2) * (_tw + 0.3), TOP + (_i // 2) * (_th + _tg)
    st.rect(s, _cx, _cy, _tw, _th, fill='surface', line='line')
    st.text(s, _cx + 0.22, _cy + 0.12, 0.8, 0.5, str(_i + 1).zfill(2), size='section_title',
            color='primary', bold=True, font='numbers')
    st.arrow(s, _cx + 1.0, _cy + 0.40, _cx + 1.28, _cy + 0.22, color='accent', lw=1.75)
    st.text(s, _cx + 0.22, _cy + 0.62, _tw - 0.44, 0.32, _head, size='card_head', color='title', bold=True,
            caps=True)
    st.text(s, _cx + 0.22, _cy + 0.94, _tw - 0.44, 0.52, _note, size='small', color='text_soft',
            spacing=1.05)
conclusion(s, 'Главное', 'Понял не «как писать код», а что именно строить дальше.', y=5.65, h=0.90)
notes(s, 'Здесь переход от личного опыта к продукту. Дальше два слайда — что именно предлагаем.')
drop_shadows(s)

# ---------------------------------------------------------------- 11. Anvil
# interaction_scheme: ромб - человек слева, Forja Anvil сверху, агент справа, модель в центре,
# результат снизу; поток идёт по кругу. Стрелки - от границы блока до границы блока, подписи сбоку.
s = st.slide(prs, 'Продукт', 'Forja Anvil: рабочее место человека и агентов',
             sub='Скачал, подключил модель, поручил задачу — настраивать нечего')
interaction_scheme(st, s, st.m, 2.20, CW, 3.35, [
    dict(name='Человек', note='ставит задачу словами', x=0.0, y=0.5, kind='party'),
    dict(name='Forja Anvil', note='чат, проекты, файлы', x=0.5, y=0.0, kind='center'),
    dict(name='Агент', note='работает в папке проекта', x=1.0, y=0.5, kind='party'),
    dict(name='Модель', note='любой провайдер из коробки', x=0.5, y=0.5, kind='system'),
    dict(name='Результат', note='принять или доработать', x=0.5, y=1.0, kind='party'),
], [(0, 1, 'задача'), (1, 2, 'поручение'), (2, 3, 'запрос'), (3, 2, 'ответ'),
    (2, 4, 'файл'), (4, 0, 'приёмка')], bw=2.3, bh=0.95)
conclusion(s, 'Суперсила', 'Свои знания + готовые помощники + любая модель без настройки',
           y=5.66, h=0.90)
st.footnote(s, 'Products/Nexa/CLAUDE.md:16-17; ТЗ агентов и наборов, 31.10.2026')
notes(s, 'Ключевое отличие от чат-ботов: результат — файл, который можно отдать руководству, '
         'а не текст в окне.')
drop_shadows(s)

# ---------------------------------------------------------------- 12. четыре слоя
# Лестница: четыре блока одной ширины, низ у всех общий (5,46), высота растёт вправо на 0,35 - верхний
# край ступеней и образует подъём. Заголовки (4,22) и пояснения (4,60) стоят у всех на одной высоте,
# поэтому строка читается поперёк, а ступени видны по верхам. В зазорах - короткие вертикальные
# стрелки от верха нижней ступени к верху следующей: подъём виден и без подписи. Верхняя ступень
# залита синим - это цель лестницы, нижние три - светло-синие, с синей кромкой сверху.
s = st.slide(prs, 'Платформа', 'Под одной фразой — четыре слоя',
             sub='Вход простой. Под ним среда, которая делает каждый следующий проект дешевле')
_FLOORS = [('Проект', 'файлы, задачи, доска, история решений'),
           ('Репозиторий', 'версии под капотом: коммит, ветка, откат'),
           ('Среда', 'агенты, навыки, правила, модели — настроено один раз'),
           ('Магазин', 'наборы помощников под профессию, в один клик')]
_FGAP, _FBASE = 0.30, 5.46                   # зазор между ступенями и общий низ лестницы
_FW = (CW - _FGAP * (len(_FLOORS) - 1)) / len(_FLOORS)
_FH = [1.45 + 0.35 * i for i in range(len(_FLOORS))]   # подъём 0,35 на ступень
for _i, (_head, _note) in enumerate(_FLOORS):
    _fx, _fy = st.m + _i * (_FW + _FGAP), _FBASE - _FH[_i]
    _fcap = _i == len(_FLOORS) - 1           # верхняя ступень - цель лестницы
    st.rect(s, _fx, _fy, _FW, _FH[_i], fill='primary' if _fcap else 'light')
    if not _fcap:                            # кромка сверху: держит верхний край ступени
        st.rect(s, _fx, _fy, _FW, 0.05, fill='primary')
    st.text(s, _fx + 0.19, 4.22, _FW - 0.38, 0.32, _head, size='card_head',
            color='white' if _fcap else 'primary', bold=True)
    st.text(s, _fx + 0.19, 4.60, _FW - 0.38, 0.68, _note, size='small',
            color='light' if _fcap else 'text', spacing=1.1)
    if _i:                                  # стрелка подъёма: от верха нижней ступени к верху этой
        _ax = _fx - _FGAP / 2
        st.arrow(s, _ax, _FBASE - _FH[_i - 1] - 0.05, _ax, _fy + 0.06, color='accent', lw=2.5)
st.line(s, st.m, _FBASE, st.m + CW, _FBASE, color='axis', lw=1.0)   # пол под лестницей
conclusion(s, 'Смысл', 'Возвращаются не за функцией, а за накопленным контекстом.', y=5.66, h=0.90)
notes(s, 'Одна фраза для всех четырёх слоёв: проект, репозиторий, среда, магазин. Чем выше ступень, '
         'тем реже человек её настраивает - и тем дороже она стоит.')
drop_shadows(s)

# ---------------------------------------------------------------- 13. бизнес-модель
# Воронка: шесть этапов полосами и плашка результата справа (см. funnel_bands - почему не st.funnel).
# Подписи этапов - глаголами: «скачал → подключил → сделал → взял → купил → открыл».
s = st.slide(prs, 'Бизнес-модель', 'Вход бесплатный, платят за помощников',
             sub='Не подписка на приложение — набор инструментов одной студии')
funnel_bands(s, st.m, TOP, CW, 3.22, [
    ('Скачал', 'Бесплатно, без терминала'),
    ('Подключил модель', 'Мастер: ключ из коробки'),
    ('Сделал проект', 'За пять минут'),
    ('Взял помощников', 'Платно: три товара'),
    ('Купил набор под профессию', 'Проектный офис, аналитик'),
    ('Открыл скины и ранг', 'Игровой слой, не за деньги'),
], result=('Результат', 'Платят за помощников и наборы, а не за приложение'))
conclusion(s, 'Следующий шаг',
           'Дальше — свой сервер: подключение моделей из коробки, без настройки провайдеров.',
           y=5.55)
st.footnote(s, 'Products/Nexa/CLAUDE.md:19-21; ТЗ 01.11.2026 (три товара, игровой ранг, юридическая рамка)')
notes(s, 'Не выдавать гипотезу за расчёт. Честно: цены взяты из проработки AgentKit, рынок не мерил.')
drop_shadows(s)

# ---------------------------------------------------------------- 14. итоги и планы
# done_plans: две колонки - слева «сделано за курс» с галочками, справа шаги со сроком.
# Колонка «сделано» начинается с 2,12: строка шапки с одной строкой подзаголовка кончается на 1,99.
s = st.slide(prs, 'Что дальше', 'Итоги и планы', sub='Что получилось и куда идём')
done_plans(st, s, st.m, 2.12, CW, 3.50,
           [('Паки агентов', 'Автоматизировали рутинные задачи: чтение, поиск, документы, проверка'),
            ('Среда и слой данных', 'Библиотека знаний, правила, хуки; смена модели за день'),
            ('Продукты', 'CoreDesk и Forja, TayaBot, доска, сайты, бот')],
           [('Ноябрь', 'Forja Anvil: бета для пяти аналитиков'),
            ('Декабрь', 'Forja Board: задача уходит агенту'),
            ('2027', 'Пять пилотов AgentKit')])
conclusion(s, 'Смысл', 'Следующий проект — не с нуля.', y=5.68, h=0.90)
st.footnote(s, 'Репозиторий: github.com/khsergey-afk/forja-zerocoder')
notes(s, 'Репозиторий: доклад, стиль, скилл, README — показать на экране. Планы — конкретные, '
         'с одним сроком.')
drop_shadows(s)

# ---------------------------------------------------------------- 15. спасибо
# Титул стиля на тёмной подложке: в нейтральном оформлении метода thanks нет, поэтому собираем
# закрывающий слайд теми же примитивами, что и титул.
s = st.blank(prs, bg='deep')
st.rect(s, 0, 0, 0.14, st.h, fill='primary')
st.text(s, st.m, 2.55, CW, 0.9, 'Спасибо за внимание', size='cover_title', color='on_dark', bold=True,
        font='heading', spacing=0.95)
st.rect(s, st.m, 3.58, 0.5, 0.03, fill='accent')
st.text(s, st.m, 3.80, CW, 0.5, 'Вопросы — с удовольствием', size='cover_subtitle', color='light')
st.text(s, st.w - st.m - 6.0, 6.25, 6.0, 0.3, 'Хачатрян Сергей Арамович, Руководитель проектов',
        size='small', color='light', align=PP_ALIGN.RIGHT)
notes(s, 'Готов ответить про слой данных, про переезд с провайдера и про экономику продукта.')
drop_shadows(s)

prs.save(str(OUT))
print(f'готово: {OUT}  слайдов: {len(prs.slides)}')
