#!/usr/bin/env python3
"""
Generate PGP for Outlook add-in icons following MS Office Add-in icon guidelines.
https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-icons

Variants:
  Icon*        - PGP Encryption group button (envelope + paper + blue padlock + crossed keys)
  IconEncrypt* - Encrypt action             (green padlock, closed)
  IconDecrypt* - Decrypt action             (red padlock,   open)
  IconKeys*    - Manage Keys button         (person silhouette + crossed keys)
"""

from PIL import Image, ImageDraw
import math
import os

# ── MS-approved colours ───────────────────────────────────────────────────────
DARK_GRAY   = (58,  58,  56)    # #3A3A38  outline / standalone
MEDIUM_GRAY = (121, 119, 116)   # #797774  secondary content
BG_FILL     = (250, 250, 250)   # #FAFAFA  paper background
LIGHT_GRAY  = (200, 198, 196)   # #C8C6C4  envelope fill

#            standalone          outline            fill
BLUE  = ((30,  139, 205), (0,   99,  177), (131, 190, 236))  # general / keys
RED   = ((237,  61,  59), (212,  35,  20), (255, 145, 152))  # decrypt
GREEN = ((24,  171,  80), (48,  144,  72), (161, 221, 170))  # encrypt
YELLOW = ((251, 152,  59), (237, 135,  51), (248, 219, 143)) # second key in pair


def draw_icon(size: int, lock_sa, lock_ol, lock_fill,
              lock_open: bool = False) -> Image.Image:
    """
    Draw the PGP Encrypt/Decrypt icon at `size` pixels square.

    Design coordinate space: 80 x 80 units.
    Supersampled 4x for smooth anti-aliasing, then downscaled with LANCZOS.

    Layout:
      * Letter/paper sheet - peeking above the envelope (>=32 px)
      * Envelope body with V-flap
      * Padlock overlay (bottom-right quadrant)
        - Closed (green / blue): centred U-shackle, both legs in body
        - Open   (red):          shackle swung left, right leg only in body
    """
    SS = 4
    W = H = size * SS
    img  = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    def p(v):
        """80-unit design coord -> supersampled pixel."""
        return round(v * W / 80)

    def lw(v):
        """Line width, minimum 1 real pixel (= SS supersampled pixels)."""
        return max(SS, round(v * W / 80))

    def rgba(c):
        return (*c, 255)

    def rounded_rect(coords, radius, fill, outline, width):
        try:
            draw.rounded_rectangle(coords, radius=radius,
                                   fill=fill, outline=outline, width=width)
        except AttributeError:
            draw.rectangle(coords, fill=fill, outline=outline, width=width)

    # ── 16 px: simplified (two elements only) ────────────────────────────────
    if size <= 16:
        # Envelope
        draw.rectangle([p(2), p(18), p(60), p(70)],
                       fill=rgba(LIGHT_GRAY), outline=rgba(DARK_GRAY),
                       width=lw(4))
        mx = p(31)
        draw.line([(p(2),  p(18)), (mx, p(36))], fill=rgba(DARK_GRAY), width=lw(4))
        draw.line([(mx, p(36)), (p(60), p(18))], fill=rgba(DARK_GRAY), width=lw(4))

        # Padlock body (bottom-right badge)
        lk_l, lk_t, lk_r, lk_b = p(46), p(44), p(76), p(76)
        rounded_rect([lk_l, lk_t, lk_r, lk_b], radius=p(4),
                     fill=rgba(lock_fill), outline=rgba(lock_ol), width=lw(4))

        slw = lw(5)
        if lock_open:
            # Shackle: arc extending left of body, right leg only
            s_l, s_t, s_r, s_b = p(26), p(22), p(62), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(3))],
                      fill=rgba(lock_ol), width=slw)
        else:
            # Shackle: centred arc, both legs
            s_l, s_t, s_r, s_b = p(50), p(22), p(72), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            draw.line([(s_l, arc_mid_y), (s_l, lk_t + lw(3))],
                      fill=rgba(lock_ol), width=slw)
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(3))],
                      fill=rgba(lock_ol), width=slw)

    # ── 32 px and above: full detailed design ────────────────────────────────
    else:
        # Paper / letter sheet (behind envelope, peeking above flap)
        draw.rectangle([p(7), p(6), p(47), p(50)],
                       fill=rgba(BG_FILL), outline=rgba(DARK_GRAY), width=lw(1.5))

        # Illegible text lines on paper
        for ly in [17, 24, 31, 38]:
            draw.line([(p(11), p(ly)), (p(43), p(ly))],
                      fill=rgba(MEDIUM_GRAY), width=lw(1.5))

        # Envelope body
        draw.rectangle([p(2), p(28), p(54), p(64)],
                       fill=rgba(LIGHT_GRAY), outline=rgba(DARK_GRAY), width=lw(1.5))

        # Envelope flap (V-shape)
        mx, fy = p(28), p(40)
        draw.line([(p(2),  p(28)), (mx, fy)], fill=rgba(DARK_GRAY), width=lw(1.5))
        draw.line([(mx, fy), (p(54), p(28))], fill=rgba(DARK_GRAY), width=lw(1.5))

        # Padlock body
        lk_l, lk_t, lk_r, lk_b = p(45), p(42), p(78), p(76)
        rounded_rect([lk_l, lk_t, lk_r, lk_b], radius=p(4),
                     fill=rgba(lock_fill), outline=rgba(lock_ol), width=lw(2))

        # Padlock shackle
        slw = lw(3)
        if lock_open:
            # Open: arc extends LEFT of padlock body; right leg only enters body
            s_l, s_t, s_r, s_b = p(32), p(22), p(64), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            # Right leg into body
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(2))],
                      fill=rgba(lock_ol), width=slw)
        else:
            # Closed: centred arc, both legs enter padlock body
            s_l, s_t, s_r, s_b = p(51), p(26), p(72), p(50)
            draw.arc([s_l, s_t, s_r, s_b],
                     start=180, end=360, fill=rgba(lock_ol), width=slw)
            arc_mid_y = (s_t + s_b) // 2
            draw.line([(s_l, arc_mid_y), (s_l, lk_t + lw(2))],
                      fill=rgba(lock_ol), width=slw)
            draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(2))],
                      fill=rgba(lock_ol), width=slw)

        # Keyhole circle on padlock face
        kh_cx = (lk_l + lk_r) // 2
        kh_cy = (lk_t + lk_b) // 2
        kh_r  = p(5)
        draw.ellipse([kh_cx - kh_r, kh_cy - kh_r,
                      kh_cx + kh_r, kh_cy + kh_r],
                     fill=rgba(lock_sa))

    return img.resize((size, size), Image.LANCZOS)


# ── Crossed-keys helpers ──────────────────────────────────────────────────────

def _draw_key_shaft(draw, hx, hy, tip_x, tip_y, colors, stroke_w, head_r,
                    draw_teeth=True):
    """Draw a key shaft from just outside the head circle to the tip, with teeth.

    hx, hy   : head centre (supersampled coords)
    tip_x/y  : tip end (supersampled coords)
    colors   : (standalone, outline, fill) tuple
    stroke_w : line width in supersampled pixels
    head_r   : head radius (to start shaft outside the circle)
    draw_teeth: whether to add teeth notches on the shaft
    """
    sa, ol, fill = colors
    rgba_ol = (*ol, 255)
    rgba_fill = (*fill, 255)

    dx = tip_x - hx
    dy = tip_y - hy
    length = math.sqrt(dx * dx + dy * dy)
    if length == 0:
        return
    ux, uy = dx / length, dy / length   # unit vector head->tip
    # Perpendicular (rotated 90 deg clockwise)
    px, py = -uy, ux

    # Shaft start: just outside head circle
    sx = hx + ux * head_r * 1.1
    sy = hy + uy * head_r * 1.1

    # Draw shaft outline then fill
    draw.line([(sx, sy), (tip_x, tip_y)], fill=rgba_ol, width=stroke_w + 2)
    draw.line([(sx, sy), (tip_x, tip_y)], fill=rgba_fill, width=stroke_w - 2)

    if draw_teeth:
        # Two teeth, each is a small rectangle notch on the perpendicular side
        tooth_w = stroke_w + 2
        tooth_h = int(stroke_w * 1.3)
        for frac in (0.55, 0.72):
            tx = hx + ux * length * frac
            ty = hy + uy * length * frac
            # Notch outward (same side for both teeth)
            notch_pts = [
                (tx + px * tooth_h,            ty + py * tooth_h),
                (tx + px * tooth_h + ux * tooth_w, ty + py * tooth_h + uy * tooth_w),
                (tx + ux * tooth_w,            ty + uy * tooth_w),
                (tx,                           ty),
            ]
            draw.polygon(notch_pts, fill=rgba_fill, outline=rgba_ol)


def _draw_key_head(draw, hx, hy, colors, head_r, stroke_w):
    """Draw a key head: filled circle with a small hole."""
    sa, ol, fill = colors
    draw.ellipse([hx - head_r, hy - head_r, hx + head_r, hy + head_r],
                 fill=(*fill, 255), outline=(*ol, 255), width=max(2, stroke_w // 2))
    # Keyhole dot
    hole_r = max(2, head_r // 3)
    draw.ellipse([hx - hole_r, hy - hole_r, hx + hole_r, hy + hole_r],
                 fill=(*sa, 255))


def _draw_crossed_keys(draw, cx, cy, key_len, colors1, colors2, stroke_w,
                        head_r, angle1_deg=45, angle2_deg=135):
    """
    Draw two crossed keys whose shafts intersect at (cx, cy).

    Key 1 (colors1): head at angle1_deg direction from center (upper-left for 45 deg)
    Key 2 (colors2): head at angle2_deg direction from center (upper-right for 135 deg)

    Z-order: key2 shaft, key1 shaft, key2 head, key1 head
    """
    HEAD_FRAC = 0.45  # head is this fraction of key_len from the crossing

    def head_pos(angle_deg):
        a = math.radians(angle_deg)
        return (cx - math.cos(a) * key_len * HEAD_FRAC,
                cy - math.sin(a) * key_len * HEAD_FRAC)

    def tip_pos(angle_deg):
        a = math.radians(angle_deg)
        return (cx + math.cos(a) * key_len * (1 - HEAD_FRAC),
                cy + math.sin(a) * key_len * (1 - HEAD_FRAC))

    h1x, h1y = head_pos(angle1_deg)
    h2x, h2y = head_pos(angle2_deg)
    t1x, t1y = tip_pos(angle1_deg)
    t2x, t2y = tip_pos(angle2_deg)

    # Z-order: key2 shaft, key1 shaft, key2 head, key1 head
    _draw_key_shaft(draw, h2x, h2y, t2x, t2y, colors2, stroke_w, head_r)
    _draw_key_shaft(draw, h1x, h1y, t1x, t1y, colors1, stroke_w, head_r)
    _draw_key_head(draw, h2x, h2y, colors2, head_r, stroke_w)
    _draw_key_head(draw, h1x, h1y, colors1, head_r, stroke_w)


# ── New main PGP icon (envelope + paper + padlock + crossed keys) ─────────────

def draw_pgp_main_icon(size: int) -> Image.Image:
    """
    PGP Encryption group/button icon.

    Layout (80-unit design space):
      * Letter/paper sheet peeking above envelope (>=32 px)
      * Envelope body with V-flap (left side)
      * Blue closed padlock (bottom-right)
      * Crossed blue+yellow keys on padlock face (>=32 px)
      * Simplified at 16 px: just envelope + padlock badge
    """
    SS = 4
    W = H = size * SS
    img  = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    lock_sa, lock_ol, lock_fill = BLUE

    def p(v):
        return round(v * W / 80)

    def lw(v):
        return max(SS, round(v * W / 80))

    def rgba(c):
        return (*c, 255)

    def rounded_rect(coords, radius, fill, outline, width):
        try:
            draw.rounded_rectangle(coords, radius=radius,
                                   fill=fill, outline=outline, width=width)
        except AttributeError:
            draw.rectangle(coords, fill=fill, outline=outline, width=width)

    if size <= 16:
        # Simplified: envelope + padlock badge only
        draw.rectangle([p(2), p(18), p(60), p(70)],
                       fill=rgba(LIGHT_GRAY), outline=rgba(DARK_GRAY), width=lw(4))
        mx = p(31)
        draw.line([(p(2), p(18)), (mx, p(36))], fill=rgba(DARK_GRAY), width=lw(4))
        draw.line([(mx, p(36)), (p(60), p(18))], fill=rgba(DARK_GRAY), width=lw(4))

        lk_l, lk_t, lk_r, lk_b = p(46), p(44), p(76), p(76)
        rounded_rect([lk_l, lk_t, lk_r, lk_b], radius=p(4),
                     fill=rgba(lock_fill), outline=rgba(lock_ol), width=lw(4))
        s_l, s_t, s_r, s_b = p(50), p(22), p(72), p(50)
        slw = lw(5)
        draw.arc([s_l, s_t, s_r, s_b], start=180, end=360,
                 fill=rgba(lock_ol), width=slw)
        arc_mid_y = (s_t + s_b) // 2
        draw.line([(s_l, arc_mid_y), (s_l, lk_t + lw(3))],
                  fill=rgba(lock_ol), width=slw)
        draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(3))],
                  fill=rgba(lock_ol), width=slw)

    else:
        # Paper sheet (behind envelope)
        draw.rectangle([p(7), p(6), p(47), p(50)],
                       fill=rgba(BG_FILL), outline=rgba(DARK_GRAY), width=lw(1.5))
        for ly in [17, 24, 31, 38]:
            draw.line([(p(11), p(ly)), (p(43), p(ly))],
                      fill=rgba(MEDIUM_GRAY), width=lw(1.5))

        # Envelope body
        draw.rectangle([p(2), p(28), p(54), p(64)],
                       fill=rgba(LIGHT_GRAY), outline=rgba(DARK_GRAY), width=lw(1.5))
        mx, fy = p(28), p(40)
        draw.line([(p(2), p(28)), (mx, fy)], fill=rgba(DARK_GRAY), width=lw(1.5))
        draw.line([(mx, fy), (p(54), p(28))], fill=rgba(DARK_GRAY), width=lw(1.5))

        # Padlock body
        lk_l, lk_t, lk_r, lk_b = p(45), p(42), p(78), p(76)
        rounded_rect([lk_l, lk_t, lk_r, lk_b], radius=p(4),
                     fill=rgba(lock_fill), outline=rgba(lock_ol), width=lw(2))

        # Padlock shackle (closed)
        slw = lw(3)
        s_l, s_t, s_r, s_b = p(51), p(26), p(72), p(50)
        draw.arc([s_l, s_t, s_r, s_b], start=180, end=360,
                 fill=rgba(lock_ol), width=slw)
        arc_mid_y = (s_t + s_b) // 2
        draw.line([(s_l, arc_mid_y), (s_l, lk_t + lw(2))],
                  fill=rgba(lock_ol), width=slw)
        draw.line([(s_r, arc_mid_y), (s_r, lk_t + lw(2))],
                  fill=rgba(lock_ol), width=slw)

        # Crossed keys on padlock face (replacing keyhole)
        kh_cx = (lk_l + lk_r) // 2
        kh_cy = (lk_t + lk_b) // 2
        key_len = p(18)
        head_r  = p(5)
        key_sw  = max(SS, p(2.5))
        _draw_crossed_keys(draw, kh_cx, kh_cy, key_len,
                           BLUE, YELLOW, key_sw, head_r,
                           angle1_deg=45, angle2_deg=135)

    return img.resize((size, size), Image.LANCZOS)


# ── Person + keys icon (Manage Keys button) ────────────────────────────────────

def draw_person_keys_icon(size: int) -> Image.Image:
    """
    Manage Keys button icon.

    Layout (80-unit design space):
      * Person silhouette: head circle + rounded body rectangle
      * Crossed blue+yellow keys overlaid on body
      * Keys-only at 16 px (person is too small to render clearly)
    """
    SS = 4
    W = H = size * SS
    img  = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    def p(v):
        return round(v * W / 80)

    def lw(v):
        return max(SS, round(v * W / 80))

    def rgba(c):
        return (*c, 255)

    def rounded_rect(coords, radius, fill, outline, width):
        try:
            draw.rounded_rectangle(coords, radius=radius,
                                   fill=fill, outline=outline, width=width)
        except AttributeError:
            draw.rectangle(coords, fill=fill, outline=outline, width=width)

    if size <= 16:
        # At 16 px the person is unreadable — draw keys only, centered
        cx, cy = p(40), p(44)
        key_len = p(30)
        head_r  = p(7)
        key_sw  = max(SS, p(3.5))
        _draw_crossed_keys(draw, cx, cy, key_len,
                           BLUE, YELLOW, key_sw, head_r,
                           angle1_deg=45, angle2_deg=135)
    else:
        person_color = MEDIUM_GRAY

        # Head circle
        hd_cx, hd_cy, hd_r = p(40), p(16), p(11)
        draw.ellipse([hd_cx - hd_r, hd_cy - hd_r, hd_cx + hd_r, hd_cy + hd_r],
                     fill=rgba(person_color), outline=rgba(DARK_GRAY), width=lw(1.5))

        # Body rounded-rect
        rounded_rect([p(10), p(30), p(70), p(74)], radius=p(8),
                     fill=rgba(person_color), outline=rgba(DARK_GRAY), width=lw(1.5))

        # Crossed keys centered on body
        cx, cy = p(40), p(54)
        key_len = p(32)
        head_r  = p(7)
        key_sw  = max(SS, p(2.5))
        _draw_crossed_keys(draw, cx, cy, key_len,
                           BLUE, YELLOW, key_sw, head_r,
                           angle1_deg=45, angle2_deg=135)

        # Re-draw head on top so keys don't obscure it
        draw.ellipse([hd_cx - hd_r, hd_cy - hd_r, hd_cx + hd_r, hd_cy + hd_r],
                     fill=rgba(person_color), outline=rgba(DARK_GRAY), width=lw(1.5))

    return img.resize((size, size), Image.LANCZOS)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    out_dir = os.path.join(os.path.dirname(__file__), 'web', 'images')
    os.makedirs(out_dir, exist_ok=True)

    # Encrypt / Decrypt icons (unchanged design)
    variants = [
        ('IconEncrypt', GREEN, False, [16, 32, 80]),
        ('IconDecrypt', RED,   True,  [16, 32, 80]),
    ]
    for prefix, colours, open_lock, sizes in variants:
        sa, ol, fill = colours
        for sz in sizes:
            img  = draw_icon(sz, sa, ol, fill, open_lock)
            path = os.path.join(out_dir, f'{prefix}{sz}.png')
            img.save(path)
            print(f'  saved  {path}')

    # Main PGP group icon (envelope + padlock + crossed keys)
    for sz in [16, 32, 64, 80, 128, 192]:
        img  = draw_pgp_main_icon(sz)
        path = os.path.join(out_dir, f'Icon{sz}.png')
        img.save(path)
        print(f'  saved  {path}')

    # Manage Keys icon (person + crossed keys)
    for sz in [16, 32, 80]:
        img  = draw_person_keys_icon(sz)
        path = os.path.join(out_dir, f'IconKeys{sz}.png')
        img.save(path)
        print(f'  saved  {path}')

    print('Done.')


if __name__ == '__main__':
    main()
