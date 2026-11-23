"""Gera o manual (HTML, visual LAWgico dos outros manuais do painel) e a apresentação (PPTX, template oficial P&C)
de cada plataforma em conteudo.py. Saída: <repo>/<pasta>/<arquivo>.html e .pptx.

Uso: python gerar_manuais.py
"""
import datetime as dt, html, pathlib, sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

AQUI = pathlib.Path(__file__).resolve().parent
REPO = AQUI.parent
sys.path.insert(0, str(AQUI))
from conteudo import PLATAFORMAS  # noqa: E402

TEMPLATE = REPO.parent / "Base diária de atualização sócios" / "Apresentação - Novo template.pptx"
HOJE = dt.date.today().strftime("%d/%m/%Y")
e = html.escape

CSS = """
:root{--navy:#051520;--navy2:#0a2235;--cyan:#00A9E0;--text:#e8edf3;--muted:#8fabc2;--border:rgba(0,169,224,.18);--card:rgba(255,255,255,.04);--amber:#E6A23C;--green:#27AE60}
*{margin:0;padding:0;box-sizing:border-box}html{scroll-behavior:smooth}
body{font-family:'Sora',sans-serif;background:var(--navy);color:var(--text);min-height:100vh;overflow-x:hidden;font-size:14px}
body::before{content:'';position:fixed;inset:0;z-index:0;pointer-events:none;background:radial-gradient(ellipse 70% 50% at 10% 0%,rgba(0,169,224,.10) 0%,transparent 60%),radial-gradient(ellipse 50% 40% at 90% 100%,rgba(0,59,92,.18) 0%,transparent 55%)}
.header{position:sticky;top:0;z-index:10;padding:18px 48px;border-bottom:1px solid var(--border);background:rgba(5,21,32,.88);backdrop-filter:blur(12px);display:flex;align-items:center;justify-content:space-between;gap:24px}
.lw{display:flex;flex-direction:column;line-height:1}.lw-w{display:flex;align-items:flex-end}.lw-law{font-size:26px;font-weight:800;color:#fff;letter-spacing:-1px}
.lw-col{display:flex;flex-direction:column}.lw-mod{font-size:9px;color:rgba(255,255,255,.75);letter-spacing:.8px;align-self:flex-end}.lw-gico{font-size:26px;font-weight:300;color:#fff;letter-spacing:-1px;line-height:1}
.lw-tag{font-size:8px;color:rgba(255,255,255,.6);letter-spacing:3px;margin-top:2px}
.hr{text-align:right}.firm{font-size:12px;font-weight:600;color:#fff;opacity:.85}.firm-sub{font-size:10px;color:var(--muted);margin-top:2px}
.badge{display:inline-block;margin-top:8px;background:rgba(0,169,224,.15);border:1px solid var(--border);color:var(--cyan);font-size:10px;font-weight:700;padding:3px 12px;border-radius:20px;letter-spacing:.5px}
.btn{display:inline-flex;align-items:center;gap:6px;margin-top:22px;background:var(--cyan);color:var(--navy);font-size:12px;font-weight:700;padding:9px 18px;border-radius:9px;text-decoration:none}
.hero{position:relative;z-index:5;padding:56px 48px 40px;border-bottom:1px solid var(--border)}
.tag{display:inline-flex;align-items:center;gap:8px;background:rgba(0,169,224,.10);border:1px solid rgba(0,169,224,.25);color:var(--cyan);font-size:10px;font-weight:700;padding:5px 16px;border-radius:20px;letter-spacing:1px;text-transform:uppercase;margin-bottom:20px}
.tag::before{content:'';width:6px;height:6px;border-radius:50%;background:var(--cyan)}
.hero h1{font-size:38px;font-weight:800;color:#fff;line-height:1.12;letter-spacing:-.5px;margin-bottom:12px;overflow-wrap:anywhere}.hero h1 span{color:var(--cyan)}
.hero p{font-size:15px;color:var(--muted);line-height:1.7;max-width:720px}
.toc{position:relative;z-index:5;padding:26px 48px;border-bottom:1px solid var(--border);background:var(--card)}
.toc-t{font-size:10px;font-weight:700;color:var(--cyan);text-transform:uppercase;letter-spacing:2px;margin-bottom:14px}
.toc-g{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px}
.toc-g a{display:flex;align-items:center;gap:10px;padding:10px 14px;border-radius:10px;background:rgba(255,255,255,.03);border:1px solid var(--border);text-decoration:none;color:var(--muted);font-size:12.5px;font-weight:500}
.toc-g a:hover{background:rgba(0,169,224,.10);border-color:var(--cyan);color:var(--cyan)}.toc-g b{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--cyan);width:18px}
.content{position:relative;z-index:5;max-width:940px;margin:0 auto;padding:0 48px 80px}
.sec{margin-top:54px;scroll-margin-top:96px}.sec-h{display:flex;align-items:center;gap:16px;margin-bottom:22px}
.sec-n{width:36px;height:36px;border-radius:10px;flex-shrink:0;background:rgba(0,169,224,.15);border:1px solid rgba(0,169,224,.3);display:flex;align-items:center;justify-content:center;font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;color:var(--cyan)}
.sec-t{font-size:21px;font-weight:700;color:#fff;letter-spacing:-.3px}.sec-l{flex:1;height:1px;background:linear-gradient(90deg,var(--border),transparent)}
p{font-size:14px;color:var(--muted);line-height:1.8;margin-bottom:14px}
ul.l{list-style:none;margin:0 0 16px}ul.l li{position:relative;padding:7px 0 7px 22px;color:var(--text);line-height:1.6;border-bottom:1px solid rgba(255,255,255,.04)}
ul.l li::before{content:'';position:absolute;left:4px;top:15px;width:7px;height:7px;border-radius:2px;background:var(--cyan)}
.card{background:var(--card);border:1px solid var(--border);border-left:3px solid var(--cyan);border-radius:14px;padding:16px 20px;margin:0 0 16px}
.card.amber{border-left-color:var(--amber)}.card h4{font-size:13.5px;color:#fff;margin-bottom:6px}.card p{margin:0}
.passos{display:grid;gap:10px;margin-bottom:16px}.passo{display:grid;grid-template-columns:34px 1fr;gap:14px;background:var(--card);border:1px solid var(--border);border-radius:12px;padding:14px 16px}
.passo i{font-style:normal;width:30px;height:30px;border-radius:50%;background:var(--cyan);color:var(--navy);font-weight:800;display:flex;align-items:center;justify-content:center}
.passo h4{font-size:14px;color:#fff;margin-bottom:4px}.passo p{margin:0;font-size:13.5px}
.tb{overflow-x:auto;margin-bottom:16px;border:1px solid var(--border);border-radius:12px}
table{width:100%;border-collapse:collapse;font-size:13px}th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--cyan);padding:11px 14px;background:rgba(0,169,224,.08)}
td{padding:11px 14px;border-top:1px solid rgba(255,255,255,.06);color:var(--muted);vertical-align:top;line-height:1.55}td:first-child{color:#fff;font-weight:600;white-space:nowrap}
footer{position:relative;z-index:5;text-align:center;padding:28px;color:var(--muted);font-size:11.5px;border-top:1px solid var(--border)}
@media (max-width:700px){.header,.hero,.toc{padding-left:16px;padding-right:16px}.content{padding:0 16px 60px}.hero h1{font-size:28px}.hr{display:none}td:first-child{white-space:normal}}
@media print{.header{position:static}body{background:#fff;color:#000}}
"""


def bloco_html(b):
    t = b[0]
    if t == "p":
        return f"<p>{e(b[1])}</p>"
    if t == "lista":
        return '<ul class="l">' + "".join(f"<li>{e(i)}</li>" for i in b[1]) + "</ul>"
    if t == "card":
        return f'<div class="card {b[1]}"><h4>{e(b[2])}</h4><p>{e(b[3])}</p></div>'
    if t == "passos":
        return '<div class="passos">' + "".join(f'<div class="passo"><i>{k}</i><div><h4>{e(a)}</h4><p>{e(d)}</p></div></div>' for k, (a, d) in enumerate(b[1], 1)) + "</div>"
    if t == "tabela":
        return ('<div class="tb"><table><thead><tr>' + "".join(f"<th>{e(h)}</th>" for h in b[1]) + "</tr></thead><tbody>" +
                "".join("<tr>" + "".join(f"<td>{e(c)}</td>" for c in r) + "</tr>" for r in b[2]) + "</tbody></table></div>")
    raise ValueError(t)


def manual_html(P):
    secs = P["secoes"]
    toc = "".join(f'<a href="#s{k}"><b>{k:02d}</b>{e(s[0])}</a>' for k, s in enumerate(secs, 1))
    corpo = "".join(f'<section class="sec" id="s{k}"><div class="sec-h"><div class="sec-n">{k:02d}</div><div class="sec-t">{e(s[0])}</div><div class="sec-l"></div></div>'
                    + "".join(bloco_html(b) for b in s[2]) + "</section>" for k, s in enumerate(secs, 1))
    titulo = f"{P['nome']} {P['destaque']}"
    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Manual {e(titulo)} | Peixoto &amp; Cury</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=IBM+Plex+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>{CSS}</style></head><body>
<header class="header"><div class="lw"><div class="lw-w"><span class="lw-law">LAW</span><div class="lw-col"><span class="lw-mod">manual</span><span class="lw-gico">gico</span></div></div><div class="lw-tag">lawyers at work</div></div>
<div class="hr"><div class="firm">Peixoto &amp; Cury Advogados</div><div class="firm-sub">Desenvolvido pelo Peixoto &amp; Cury</div><span class="badge">Manual do usuário · {HOJE}</span></div></header>
<section class="hero"><span class="tag">Manual do usuário</span><h1>{e(P['nome'])} <span>{e(P['destaque'])}</span></h1><p>{e(P['sub'])}</p>
<a class="btn" href="../" target="_top">Abrir pelo painel LAWgico →</a></section>
<nav class="toc"><div class="toc-t">Neste manual</div><div class="toc-g">{toc}</div></nav>
<main class="content">{corpo}</main>
<footer>Peixoto &amp; Cury Advogados · LAWgico · manual atualizado em {HOJE}</footer>
</body></html>"""


# ------------------------------------------------------------------ PPTX
NAVY, CYAN, TEXTO, BRANCO, SUAVE, BORDA, AMBAR = "0E2841", "009DDC", "33475B", "FFFFFF", "5B6B7D", "D5DCE8", "E6A23C"
FONTE = "Trebuchet MS"
cor = RGBColor.from_string


def caixa_txt(sl, x, y, w, h, paras, size=16, color=TEXTO, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, espaco=4):
    cx = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = cx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for lado in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, lado, Inches(0.05))
    paras = paras if isinstance(paras, list) else [paras]
    for k, par in enumerate(paras):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(espaco)
        runs = par if isinstance(par, list) else [(par, {})]
        for txt, o in runs:
            r = p.add_run()
            r.text = txt
            r.font.name = FONTE
            r.font.size = Pt(o.get("size", size))
            r.font.bold = o.get("bold", bold)
            r.font.color.rgb = cor(o.get("color", color))
    return cx


def retangulo(sl, x, y, w, h, fill=BRANCO, line=BORDA, forma=MSO_SHAPE.ROUNDED_RECTANGLE):
    s = sl.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = cor(fill)
    if line:
        s.line.color.rgb = cor(line)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    if forma == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def selo(sl, x, y, rot, d=0.55):
    c = retangulo(sl, x, y, d, d, fill=NAVY, line=None, forma=MSO_SHAPE.OVAL)
    tf = c.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = rot
    r.font.name, r.font.size, r.font.bold = FONTE, Pt(15), True
    r.font.color.rgb = cor(CYAN)


def apresentacao(P, saida):
    prs = Presentation(str(TEMPLATE))
    lista = prs.slides._sldIdLst
    for sid in list(lista):
        prs.part.drop_rel(sid.rId)
        lista.remove(sid)
    lay = {str(l.part.partname).rsplit("/", 1)[-1]: l for m in prs.slide_masters for l in m.slide_layouts}
    CAPA, INTERNO, FIM = (lay[f"slideLayout{i}.xml"] for i in (1, 6, 8))
    titulo_plat = f"{P['nome']} {P['destaque']}"
    n = [1]

    def interno(tit, sub):
        sl = prs.slides.add_slide(INTERNO)
        n[0] += 1
        caixa_txt(sl, 0.6, 0.42, 10.4, 0.7, tit, size=28, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        if sub:
            caixa_txt(sl, 0.6, 1.1, 10.4, 0.45, sub, size=15, color=SUAVE)
        caixa_txt(sl, 0.6, 6.86, 9, 0.35, f"{titulo_plat}  |  Manual do usuário", size=11, color=BRANCO, anchor=MSO_ANCHOR.MIDDLE)
        caixa_txt(sl, 11.9, 6.86, 0.85, 0.35, str(n[0]), size=11, color=BRANCO, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        return sl

    sl = prs.slides.add_slide(CAPA)
    caixa_txt(sl, 0.75, 2.55, 8.7, 1.5, [[(P["nome"] + " " + P["destaque"], {})], [("Manual do usuário", {"size": 18, "bold": False})]],
              size=30, color=NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    caixa_txt(sl, 0.75, 4.17, 6.6, 0.48, "LAWgico  |  Peixoto & Cury Advogados", size=18, color=BRANCO, anchor=MSO_ANCHOR.MIDDLE)
    caixa_txt(sl, 0.75, 5.0, 8, 0.4, f"Atualizado em {HOJE}", size=14, color=SUAVE)

    # roteiro
    sl = interno("Neste manual", P["sub"])
    secs = P["secoes"]
    for k, s in enumerate(secs):
        col, lin = k % 2, k // 2
        x, y = 0.6 + col * 6.15, 1.85 + lin * 1.45
        retangulo(sl, x, y, 5.9, 1.25)
        selo(sl, x + 0.25, y + 0.35, f"{k + 1:02d}")
        caixa_txt(sl, x + 1.0, y + 0.18, 4.7, 0.95, [[(s[0], {"bold": True, "color": NAVY, "size": 17})], [(s[1], {"size": 12.5, "color": SUAVE})]], espaco=2)

    for k, (tit, resumo, blocos) in enumerate(secs, 1):
        sl = interno(f"{k:02d}  {tit}", resumo)
        y = 1.8
        for b in blocos:
            if y > 6.2:
                break
            t = b[0]
            if t == "p":
                h = 0.32 * (len(b[1]) // 105 + 1) + 0.15
                caixa_txt(sl, 0.6, y, 12.1, h, b[1], size=14.5)
                y += h + 0.05
            elif t == "lista":
                itens = b[1][:7]
                h = 0.4 * len(itens) + 0.1
                caixa_txt(sl, 0.6, y, 12.1, h, [[("■  ", {"color": CYAN, "size": 11}), (i, {})] for i in itens], size=14.5, espaco=5)
                y += h + 0.1
            elif t == "card":
                retangulo(sl, 0.6, y, 12.1, 0.95, fill="F2F7FB", line=(AMBAR if b[1] == "amber" else CYAN))
                caixa_txt(sl, 0.85, y + 0.1, 11.6, 0.8, [[(b[2], {"bold": True, "color": NAVY, "size": 14})], [(b[3], {"size": 12.5})]], espaco=2)
                y += 1.1
            elif t == "passos":
                passos = b[1][:4]
                w = (12.1 - 0.25 * (len(passos) - 1)) / len(passos)
                for i, (a, d) in enumerate(passos):
                    x = 0.6 + i * (w + 0.25)
                    retangulo(sl, x, y, w, 2.45)
                    selo(sl, x + 0.2, y + 0.2, str(i + 1), d=0.5)
                    caixa_txt(sl, x + 0.2, y + 0.8, w - 0.4, 1.6, [[(a, {"bold": True, "color": NAVY, "size": 14})], [(d, {"size": 11})]], espaco=3)
                y += 2.65
            elif t == "tabela":
                linhas = b[2]
                rows = len(linhas) + 1
                alt = min(4.9, 0.36 * rows)
                tb = sl.shapes.add_table(rows, 2, Inches(0.6), Inches(y), Inches(12.1), Inches(alt)).table
                tb.columns[0].width = Inches(3.2)
                tb.columns[1].width = Inches(8.9)
                for i, linha in enumerate([b[1]] + linhas):
                    for j, val in enumerate(linha[:2]):
                        c = tb.cell(i, j)
                        c.fill.solid()
                        c.fill.fore_color.rgb = cor(NAVY if i == 0 else ("FFFFFF" if i % 2 else "F2F7FB"))
                        c.margin_top = c.margin_bottom = Inches(0.03)
                        p = c.text_frame.paragraphs[0]
                        p.text = ""
                        r = p.add_run()
                        r.text = val
                        r.font.name = FONTE
                        r.font.size = Pt(11 if rows > 9 else 12)
                        r.font.bold = i == 0 or j == 0
                        r.font.color.rgb = cor(BRANCO if i == 0 else (NAVY if j == 0 else TEXTO))
                y += alt + 0.15

    prs.slides.add_slide(FIM)   # encerramento do template, só com o logo
    prs.save(str(saida))


def main():
    for P in PLATAFORMAS:
        pasta = REPO / P["pasta"]
        pasta.mkdir(exist_ok=True)
        h = manual_html(P)
        assert "—" not in h and "–" not in h, P["id"]
        (pasta / f"{P['arquivo']}.html").write_text(h, encoding="utf-8")
        apresentacao(P, pasta / f"{P['arquivo']}.pptx")
        print("ok", P["id"])


if __name__ == "__main__":
    main()
