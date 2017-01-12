#! /usr/bin/env python

# http://stackoverflow.com/questions/25019441/arc-pie-cut-in-svgwrite

import sys
import math
import colorsys
import openpyxl
import svgwrite


LINE_COLOUR = 'white'
FONT_COLOUR = '#555555'
#FONT_SIZE_STR = str(10 * SCALE) + 'px'
FONT_SIZE = 12
FONT_SIZE_STR = str(FONT_SIZE) + 'px'
FONT_NAME = 'Helvetica'
COLOR_LIGHTNESS = 0.30
COLOR_INCREMENT = 0.10
RING_WIDTH = 60
RING_START = 100

HUES = [
    0.0   / 255,
    70.0  / 255,
    15.0  / 255,
    145.0 / 255,
    235.0 / 255,
    120.0 / 255,
    35.0  / 255,
    210.0 / 255,
    100.0 / 255,
]

def addArc(dwg, group, ctr, iradius, oradius, t0, t1, color):
    """ Adds an arc that bulges to the right as it moves from p0 to p1 """
    if t1 < t0:
        # swap -- always do the increasing arc...
        t = t0
        t0 = t1
        t1 = t

    while t1 - t0 > 2 * math.pi:
        t1 = t1 - 2 * math.pi

    t2 = None
    if t1 - t0 > math.pi:
        t2 = t1
        t1 = (t0 + t2) / 2

    args = {
        'x0i'      : ctr[0] + iradius * math.cos(t0),
        'y0i'      : ctr[1] - iradius * math.sin(t0),
        'x1i'      : ctr[0] + iradius * math.cos(t1),
        'y1i'      : ctr[1] - iradius * math.sin(t1),
        'x2i'      : ctr[0] + iradius * math.cos(t2) if t2 is not None else 0,
        'y2i'      : ctr[1] - iradius * math.sin(t2) if t2 is not None else 0,
        'x0o'      : ctr[0] + oradius * math.cos(t0),
        'y0o'      : ctr[1] - oradius * math.sin(t0),
        'x1o'      : ctr[0] + oradius * math.cos(t1),
        'y1o'      : ctr[1] - oradius * math.sin(t1),
        'x2o'      : ctr[0] + oradius * math.cos(t2) if t2 is not None else 0,
        'y2o'      : ctr[1] - oradius * math.sin(t2) if t2 is not None else 0,
        'xiradius' : iradius,
        'yiradius' : iradius,
        'xoradius' : oradius,
        'yoradius' : oradius,
        'rot'      : 0, # has no effect for circles
    }

    path = 'M {x0i},{y0i}'.format(**args)
    path += ' A {xiradius},{yiradius} {rot} 0,0 {x1i},{y1i}'.format(**args)
    if t2 is not None:
        path += ' A {xiradius},{yiradius} {rot} 0,0 {x2i},{y2i}'.format(**args)
        path += ' L {x2o},{y2o}'.format(**args)
        path += ' A {xoradius},{yoradius} {rot} 0,1 {x1o},{y1o}'.format(**args)
    else:
        path += ' L {x1o},{y1o}'.format(**args)
    path +=  ' A {xoradius},{yoradius} {rot} 0,1 {x0o},{y0o}'.format(**args)
    path += ' Z'
        
    group.add(dwg.path(d = path, fill = color, stroke = 'black',
                               fill_opacity = '1', stroke_width = 2))


def get_data(fn):
    wb = openpyxl.load_workbook(fn)
    ws = wb.active

    path = list()
    data = { 'children' : dict(), 'ts' : 0, 'te' : 2 * math.pi }
    for row in ws.rows:
        for idx, col in enumerate(row[:-1]):
            if col.value is not None:
                del path[idx:]
                path.append(col.value)

        cur = data
        for p in path:
            if p not in cur['children']:
                cur['children'][p] = { 'children' : dict() }
            cur = cur['children'][p]
        cur['value'] = row[-1].value

    return data


def propogate_values(data):
    if len(data['children']) == 0:
        return
    value = 0
    for key in data['children']:
        propogate_values(data['children'][key])
    for key in data['children']:
        value += data['children'][key]['value']
    data['value'] = value


def propogate_geo(data, radius):
    cv = 0
    v = data['value']
    ts = data['ts']
    te = data['te']
    for key in sorted(data['children']):
        data['children'][key]['ts'] = ts + (cv * (te - ts) / v)
        cv += data['children'][key]['value']
        data['children'][key]['te'] = ts + (cv * (te - ts) / v)
        data['children'][key]['r0'] = radius
        data['children'][key]['r1'] = radius + RING_WIDTH
        propogate_geo(data['children'][key], radius + RING_WIDTH)


def propogate_color(data, hue, lvl = 0):
    # https://www.w3.org/TR/SVG11/types.html#ColorKeywords
    rgb = colorsys.hls_to_rgb(hue, COLOR_LIGHTNESS + lvl * COLOR_INCREMENT, 1)
    data['rgb'] = tuple([ int(255*x) for x in list(rgb) ])

    for key in sorted(data['children']):
        propogate_color(data['children'][key], hue, lvl + 1)
    

def add_arcs(dwg, group, data, ctr):
    for key in data['children']:
        c = data['children'][key]
        color = 'rgb({0},{1},{2})'.format(*c['rgb'])
        addArc(dwg, group, ctr, iradius = c['r0'], oradius = c['r1'],
               t0 = c['ts'], t1 = c['te'], color = color)
        tm = (c['ts'] + c['te']) / 2
        x = ctr[0] + (((c['r0'] + c['r1']) / 2) * math.cos(tm))
        y = ctr[1] - (((c['r0'] + c['r1']) / 2) * math.sin(tm))
        a = dwg.text(key, insert = (x, y + FONT_SIZE / 2),
                     font_size = FONT_SIZE_STR,
                     text_anchor = 'middle',
                     stroke = 'none', fill = 'black')
        group.add(a)
        # group.add(dwg.line(start = (x, y-5), end = (x, y+5), stroke = 'green'))
        # group.add(dwg.line(start = (x-5, y), end = (x+5, y), stroke = 'green'))
        add_arcs(dwg, group, c, ctr)


def add_test_arcs(dwg, group):
    addArc(dwg, group, ctr = [512, 384], iradius = 250, oradius = 270,
           t0 = math.pi / 4, t1 = 3 * math.pi / 4, color = 'green')

    addArc(dwg, group, ctr = [512, 384], iradius = 270, oradius = 300,
           t0 = 0, t1 = 3 * math.pi / 2, color = 'purple')

    addArc(dwg, group, ctr = [512, 384], iradius = 300, oradius = 320,
           t0 = 0, t1 = 2 * math.pi, color = 'lavender')


def colorbar(dwg, group):
    for h in range(0,255):
        rgb = colorsys.hls_to_rgb(h/255.0, COLOR_LIGHTNESS, 0.1)
        rgb = tuple([ int(255*x) for x in list(rgb) ])
        color = 'rgb({0},{1},{2})'.format(*rgb)
        group.add(dwg.rect(insert = (4 * h, 700), size=(4, 20),
                           stroke = color, fill = color))

    for h in range(0,255):
        rgb = colorsys.hls_to_rgb(h/255.0, COLOR_LIGHTNESS, 0.5)
        rgb = tuple([ int(255*x) for x in list(rgb) ])
        color = 'rgb({0},{1},{2})'.format(*rgb)
        group.add(dwg.rect(insert = (4 * h, 720), size=(4, 20),
                           stroke = color, fill = color))

    for h in range(0,255):
        rgb = colorsys.hls_to_rgb(h/255.0, COLOR_LIGHTNESS, 0.9)
        rgb = tuple([ int(255*x) for x in list(rgb) ])
        color = 'rgb({0},{1},{2})'.format(*rgb)
        group.add(dwg.rect(insert = (4 * h, 740), size=(4, 20),
                           stroke = color, fill = color))


def main():
    if len(sys.argv) <= 1:
        print 'usage: {0} <spreadsheet>'.format(sys.argv[0])
        return

    data = get_data(sys.argv[1])
    propogate_values(data)
    propogate_geo(data, RING_START)

    N = len(data['children'])
    for idx, key in enumerate(sorted(data['children'].keys())):
        #propogate_color(data['children'][key], float(idx+1) / N+1)
        propogate_color(data['children'][key], HUES[idx])

    dwg = svgwrite.Drawing(filename="test.svg", debug=True, size=(1024,768))
    group = dwg.add(dwg.g(id = 'name', stroke = 'black', stroke_width = 1,
                          fill = 'none', fill_opacity = 1 ))
    group.add(dwg.rect(insert = (0,0), size = (1024-1,768-1)))

    add_arcs(dwg, group, data, ctr = [512, 384])

    #add_test_arcs(dwg, group)
    #colorbar(dwg, group)

    dwg.save()


if __name__ == '__main__':
    main()

