from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

SW = Inches(13.333); SH = Inches(7.5)
WHITE = RGBColor(0xFF,0xFF,0xFF)
PINK = RGBColor(0xE6,0x00,0x7E)
NAVY = RGBColor(0x1A,0x1F,0x3D)

def R(sl,l,t,w,h,f,ln=None):
    s=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,l,t,w,h)
    s.fill.solid();s.fill.fore_color.rgb=f
    if ln: s.line.fill.solid();s.line.fill.fore_color.rgb=ln;s.line.width=Pt(1)
    else: s.line.fill.background()
    return s

def T(sl,l,t,w,h,txt,sz,c,b=False,a=PP_ALIGN.LEFT,i=False):
    tb=sl.shapes.add_textbox(l,t,w,h);tb.text_frame.word_wrap=True
    p=tb.text_frame.paragraphs[0];p.text=txt;p.font.size=Pt(sz)
    p.font.color.rgb=c;p.font.bold=b;p.font.name="Arial";p.alignment=a;p.font.italic=i
    return tb

def RT(sl,l,t,w,h,runs,a=PP_ALIGN.LEFT):
    tb=sl.shapes.add_textbox(l,t,w,h);tb.text_frame.word_wrap=True
    p=tb.text_frame.paragraphs[0];p.alignment=a
    for txt,sz,c,b,i in runs:
        r=p.add_run();r.text=txt;r.font.size=Pt(sz);r.font.color.rgb=c;r.font.bold=b;r.font.name="Arial";r.font.italic=i
    return tb

def BUL(sl,l,t,w,items,sc,tc,tmpl=False,tmc=None):
    tb=sl.shapes.add_textbox(l,t,w,Inches(3.5));tf=tb.text_frame;tf.word_wrap=True
    for i,item in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_before=Pt(4);p.space_after=Pt(2)
        if tmpl:
            r=p.add_run();r.text="\u2022  "+item;r.font.size=Pt(9.5);r.font.color.rgb=tmc;r.font.italic=True;r.font.name="Arial"
        else:
            b,rest=item
            r1=p.add_run();r1.text="\u2022  "+b+" ";r1.font.size=Pt(9.5);r1.font.color.rgb=sc;r1.font.bold=True;r1.font.name="Arial"
            r2=p.add_run();r2.text=rest;r2.font.size=Pt(9.5);r2.font.color.rgb=tc;r2.font.name="Arial"

def OVL(sl,l,t,sz,n,bg):
    s=sl.shapes.add_shape(MSO_SHAPE.OVAL,l,t,sz,sz)
    s.fill.solid();s.fill.fore_color.rgb=bg;s.line.fill.background()
    p=s.text_frame.paragraphs[0];p.text=str(n);p.font.size=Pt(10);p.font.color.rgb=WHITE;p.font.bold=True;p.font.name="Arial";p.alignment=PP_ALIGN.CENTER

def bg(sl,c): sl.background.fill.solid();sl.background.fill.fore_color.rgb=c

# Data
MET=[("45%","Manual Effort Reduction"),("1.5x","Throughput Increase"),("~19.5%","Total Cost Saving"),("2 wk","Pilot to Production")]
TMET=[("___%","Headline metric"),("___x","Secondary metric"),("___","Cost / time saving"),("___","Deploy timeline")]
C1=[("Requirements are the #1 bottleneck","in SDLC \u2014 30% of projects delayed by spec gaps"),("64% of failed projects","trace root cause to errors in requirements gathering"),("Rework costs 15x more","when defects from requirements surface in production"),("BAs spend most time on low-value tasks","\u2014 formatting & consistency checks"),("No scalable alternative","\u2014 tools lack financial services domain context")]
C2=[("Automates generation","of epics, stories, and test cases from high-level inputs"),("Custom-built AI","trained on 6 corpuses of regulatory & standards docs"),("Outputs conform to internal standards","\u2014 templates, taxonomy, criteria"),("Model & tech agnostic","\u2014 swap underlying LLMs as capabilities evolve"),("Zero-disruption deployment","\u2014 parallel to existing pipelines; 2 wk pilot")]
C3=[("Up to 45% reduction","in manual effort for requirements authoring"),("1.5x throughput","\u2014 higher-quality requirements in less time"),("Accuracy & consistency","improved from first draft, fewer review cycles"),("BAs shift to high-value work","\u2014 stakeholder engagement & strategy"),("Sustainable & scalable","\u2014 pilot-to-production proven in 7 weeks")]
T1=["What is the business context?","What is the core problem?","What makes it worse?","Why isn't it solved?","Cost of inaction?"]
T2=["What does the solution do?","Key features?","What's differentiated?","How does it integrate?","How fast to deploy?"]
T3=["Headline metric?","Secondary metric?","Qualitative improvements?","How does work change?","Is it scalable?"]

L_BG=RGBColor(0xED,0xF1,0xF5);L_MID=RGBColor(0x4A,0x50,0x68);L_DIM=RGBColor(0x7B,0x81,0x98)
L_RIB=RGBColor(0xF0,0xF3,0xF8);L_DIV=RGBColor(0xD4,0xDA,0xE5);L_FT=RGBColor(0xFA,0xFB,0xFD)
COL3C=RGBColor(0x25,0x63,0xEB)

# ═══ LIGHT ═══
def light():
    prs=Presentation();prs.slide_width=SW;prs.slide_height=SH
    # Cover
    sl=prs.slides.add_slide(prs.slide_layouts[6]);bg(sl,L_BG)
    # Right navy panel
    R(sl,Emu(int(SW.emu*0.58)),Inches(0),Emu(int(SW.emu*0.42)),SH,NAVY)
    R(sl,Emu(int(SW.emu*0.58)),Inches(7.44),Emu(int(SW.emu*0.42)),Inches(0.06),PINK)
    # Left content
    T(sl,Inches(0.8),Inches(0.6),Inches(3),Inches(0.4),"CAPCO",14,L_DIM,b=True)
    R(sl,Inches(0.8),Inches(2.2),Inches(0.5),Inches(0.05),PINK)
    T(sl,Inches(0.8),Inches(2.5),Inches(6.5),Inches(1.8),"AI Demo\nOne Slider",48,NAVY,b=True)
    T(sl,Inches(0.8),Inches(4.5),Inches(6.5),Inches(0.8),"Your executive primer before the live demo \u2014 the business problem, the AI solution, and the impact, all on a single slide.",15,L_MID)
    # Format hint box
    R(sl,Inches(0.7),Inches(6.1),Inches(6.2),Inches(0.6),WHITE,ln=L_DIV)
    RT(sl,Inches(0.9),Inches(6.2),Inches(5.8),Inches(0.4),[("\u25b6 ",10,PINK,False,False),("[Solution Name]",12,NAVY,True,False),(" \u2014 Short description",12,L_DIM,False,False)])
    # Right panel content
    rx=Emu(int(SW.emu*0.58))+Inches(0.6)
    T(sl,rx,Inches(1.0),Inches(4),Inches(0.3),"SLIDE STRUCTURE",9,RGBColor(0x55,0x5A,0x78),b=True)
    for i,(num,lbl,dc) in enumerate([(1,"Problem & Complications",PINK),(2,"AI Solution & Features",RGBColor(0x88,0x8C,0xA8)),(3,"Impact & Outcomes",RGBColor(0x60,0xA5,0xFA))]):
        y=Inches(2.0+i*1.7)
        T(sl,rx,y,Inches(1),Inches(0.5),"0"+str(num),30,RGBColor(0x25,0x28,0x45),b=True)
        R(sl,rx+Inches(1.2),y+Inches(0.1),Inches(0.12),Inches(0.12),dc)
        T(sl,rx+Inches(1.5),y+Inches(0.02),Inches(3),Inches(0.4),lbl,14,RGBColor(0xE0,0xE2,0xEE),b=True)

    # Content helper
    def lcontent(title,sub,mets,c1d,c2d,c3d,tmpl=False,fr=""):
        sl2=prs.slides.add_slide(prs.slide_layouts[6]);bg(sl2,WHITE)
        R(sl2,Inches(0),Inches(0),SW,Inches(0.95),NAVY)
        R(sl2,Inches(0.6),Inches(0.22),Inches(0.05),Inches(0.5),PINK)
        T(sl2,Inches(0.85),Inches(0.2),Inches(8),Inches(0.4),title,18,WHITE,b=True)
        T(sl2,Inches(0.85),Inches(0.55),Inches(8),Inches(0.3),sub,9,RGBColor(0x99,0x9C,0xB0))
        T(sl2,Inches(10.5),Inches(0.3),Inches(2.5),Inches(0.4),"CAPCO",13,RGBColor(0x88,0x8C,0xA8),b=True,a=PP_ALIGN.RIGHT)
        rt=Inches(0.95);rh=Inches(0.85);R(sl2,Inches(0),rt,SW,rh,L_RIB)
        mw=SW.emu//4
        for i,(v,d) in enumerate(mets):
            x=Emu(mw*i);vc=PINK if i==0 and not tmpl else NAVY
            if tmpl: vc=L_DIV
            T(sl2,x+Inches(0.5),rt+Inches(0.12),Inches(1.5),Inches(0.5),v,24 if not tmpl else 20,vc,b=True)
            T(sl2,x+Inches(2.0),rt+Inches(0.18),Inches(1.5),Inches(0.5),d,8.5,L_MID if not tmpl else L_DIM,i=tmpl)
            if i<3: R(sl2,Emu(mw*(i+1)),rt+Inches(0.15),Inches(0.01),Inches(0.55),L_DIV)
        R(sl2,Inches(0),rt+rh,SW,Inches(0.01),L_DIV)
        ct=Inches(1.85);cw=SW.emu//3
        for i,(bul,lbl,acc,num) in enumerate([(c1d,"PROBLEM & COMPLICATIONS",PINK,1),(c2d,"AI SOLUTION & FEATURES",NAVY,2),(c3d,"IMPACT & OUTCOMES",COL3C,3)]):
            x=Emu(cw*i)
            OVL(sl2,x+Inches(0.5),ct+Inches(0.05),Inches(0.28),num,acc)
            T(sl2,x+Inches(0.85),ct+Inches(0.07),Inches(3),Inches(0.3),lbl,10,acc,b=True)
            R(sl2,x+Inches(0.5),ct+Inches(0.42),Inches(3.5),Inches(0.035),acc)
            BUL(sl2,x+Inches(0.5),ct+Inches(0.55),Inches(3.7),bul,NAVY,L_MID,tmpl=tmpl,tmc=L_DIM)
            if i<2: R(sl2,Emu(cw*(i+1)),ct,Inches(0.01),Inches(4.8),L_DIV)
        ft=Inches(6.95);R(sl2,Inches(0),ft,SW,Inches(0.55),L_FT);R(sl2,Inches(0),ft,SW,Inches(0.01),L_DIV)
        R(sl2,Inches(0.6),ft+Inches(0.2),Inches(0.08),Inches(0.08),PINK)
        T(sl2,Inches(0.8),ft+Inches(0.12),Inches(4),Inches(0.3),"Confidential \u2014 Capco",8,L_DIM)
        T(sl2,Inches(8),ft+Inches(0.12),Inches(5),Inches(0.3),fr,8,L_DIM,a=PP_ALIGN.RIGHT)

    lcontent("SDLC Transformation with BA Genie","Financial Services \u2022 Requirements Engineering \u2022 Custom AI Solution",MET,C1,C2,C3,fr="BA Genie \u2022 SDLC / Engineering")
    lcontent("[Solution Name] \u2014 [Use Case Title]","[Industry] \u2022 [Function / Process Area] \u2022 [Solution Type]",TMET,T1,T2,T3,tmpl=True,fr="[Solution Name] \u2022 [Domain]")
    prs.save("/home/user/Intake-to-GTM/Demo One Slider - Light.pptx");print("Light done")

# ═══ NARRATIVE ═══
AMBER=RGBColor(0xF5,0x9E,0x0B);TEAL=RGBColor(0x0E,0xA5,0xE9);EMER=RGBColor(0x10,0xB9,0x81)
D_BG=RGBColor(0x16,0x16,0x1E);D_SF=RGBColor(0x1E,0x1E,0x28);D_BD=RGBColor(0x2A,0x2A,0x38)
D_TXT=RGBColor(0xD0,0xD0,0xDC);D_MID=RGBColor(0x88,0x88,0xA0);D_DIM=RGBColor(0x50,0x50,0x68)

def narrative():
    prs=Presentation();prs.slide_width=SW;prs.slide_height=SH
    # Cover
    sl=prs.slides.add_slide(prs.slide_layouts[6]);bg(sl,D_BG)
    R(sl,Emu(int(SW.emu*0.58)),Inches(0),Emu(int(SW.emu*0.42)),SH,D_SF)
    T(sl,Inches(0.8),Inches(0.6),Inches(3),Inches(0.3),"CAPCO",14,D_DIM,b=True)
    # Colored dots
    for i,(c,xo) in enumerate([(AMBER,0),(TEAL,0.18),(EMER,0.36)]):
        R(sl,Inches(0.8+xo),Inches(1.8),Inches(0.1),Inches(0.1),c)
    R(sl,Inches(0.8),Inches(2.2),Inches(0.6),Inches(0.05),AMBER)
    T(sl,Inches(0.8),Inches(2.5),Inches(6.5),Inches(1.8),"AI Demo\nOne Slider",52,WHITE,b=True)
    T(sl,Inches(0.8),Inches(4.6),Inches(6.5),Inches(0.8),"One slide to set the stage. Walk through the problem, the AI solution, and the measurable impact.",16,D_MID)
    R(sl,Inches(0.7),Inches(6.1),Inches(6.2),Inches(0.6),D_SF,ln=D_BD)
    RT(sl,Inches(0.9),Inches(6.2),Inches(5.8),Inches(0.4),[("\u25b6 ",10,TEAL,False,False),("[Solution Name]",12,WHITE,True,False),(" \u2014 Short description",12,D_DIM,False,False)])
    # Right panel
    rx=Emu(int(SW.emu*0.58))+Inches(0.6)
    T(sl,rx,Inches(1.0),Inches(4),Inches(0.3),"NARRATIVE FLOW",9,D_DIM,b=True)
    for i,(num,lbl,ac) in enumerate([(1,"Problem & Complications",AMBER),(2,"AI Solution & Features",TEAL),(3,"Impact & Outcomes",EMER)]):
        y=Inches(2.0+i*1.7)
        T(sl,rx,y,Inches(1),Inches(0.5),"0"+str(num),30,RGBColor(0x28,0x28,0x38),b=True)
        R(sl,rx+Inches(1.2),y+Inches(0.1),Inches(0.12),Inches(0.12),ac)
        T(sl,rx+Inches(1.5),y+Inches(0.02),Inches(3),Inches(0.4),lbl,14,ac,b=True)

    # Content helper
    def ncontent(title,sub,mets,c1d,c2d,c3d,tmpl=False,fr="",stat_val="64%",stat_lbl="of failed projects\nfrom bad requirements"):
        sl2=prs.slides.add_slide(prs.slide_layouts[6]);bg(sl2,D_BG)
        R(sl2,Inches(0),Inches(0),SW,Inches(0.7),D_SF)
        T(sl2,Inches(0.6),Inches(0.1),Inches(9),Inches(0.35),title,17,WHITE if not tmpl else D_DIM,b=True)
        T(sl2,Inches(0.6),Inches(0.42),Inches(9),Inches(0.25),sub,8.5,D_DIM)
        T(sl2,Inches(10.5),Inches(0.2),Inches(2.5),Inches(0.3),"CAPCO",11,D_DIM,b=True,a=PP_ALIGN.RIGHT)
        # 3 bands
        bh=Inches(2.05);band_colors=[(AMBER,"PROBLEM"),(TEAL,"SOLUTION"),(EMER,"IMPACT")]
        for bi,(ac,lbl) in enumerate(band_colors):
            bt=Inches(0.7+bi*2.1)
            # Band bg with subtle accent
            R(sl2,Inches(0),bt,SW,bh,D_BG)
            R(sl2,Inches(0),bt,Inches(0.05),bh,ac)  # left accent
            if bi<2: R(sl2,Inches(0),bt+bh-Inches(0.01),SW,Inches(0.01),D_BD)
            # Vertical label
            T(sl2,Inches(0.15),bt+Inches(0.1),Inches(0.6),Inches(0.3),lbl,7,ac,b=True)

        # Band 1 — Problem
        bt1=Inches(0.7)
        if tmpl:
            BUL(sl2,Inches(0.8),bt1+Inches(0.4),Inches(8),c1d,WHITE,D_TXT,tmpl=True,tmc=D_DIM)
            R(sl2,Inches(10.5),bt1+Inches(0.3),Inches(2.2),Inches(1.5),D_SF,ln=D_BD)
            T(sl2,Inches(10.7),bt1+Inches(0.5),Inches(1.8),Inches(0.5),"___",28,D_BD,b=True,a=PP_ALIGN.CENTER)
            T(sl2,Inches(10.7),bt1+Inches(1.1),Inches(1.8),Inches(0.5),"Key problem\nstatistic",7.5,D_DIM,a=PP_ALIGN.CENTER,i=True)
        else:
            BUL(sl2,Inches(0.8),bt1+Inches(0.4),Inches(8),c1d,WHITE,D_TXT)
            R(sl2,Inches(10.5),bt1+Inches(0.3),Inches(2.2),Inches(1.5),D_SF,ln=D_BD)
            T(sl2,Inches(10.7),bt1+Inches(0.45),Inches(1.8),Inches(0.5),stat_val,32,AMBER,b=True,a=PP_ALIGN.CENTER)
            T(sl2,Inches(10.7),bt1+Inches(1.05),Inches(1.8),Inches(0.5),stat_lbl,7.5,D_MID,a=PP_ALIGN.CENTER)

        # Band 2 — Solution (feature chips)
        bt2=Inches(2.8)
        chip_w=Inches(2.3);chip_h=Inches(1.4);chip_gap=Inches(0.15)
        if tmpl:
            chip_data=[("[Feature 1]","What does it do?"),("[Feature 2]","What does it do?"),("[Feature 3]","What's different?"),("[Integration]","How plug in?"),("[Deploy]","Timeline?")]
        else:
            chip_data=[("Auto-Generate","Epics, stories, test cases"),("Custom AI","6 regulatory corpuses"),("Standards-Aligned","Client templates & taxonomy"),("Model Agnostic","Swap LLMs freely"),("Zero Disruption","2-week parallel pilot")]
        for ci,(ct,cd) in enumerate(chip_data):
            cx=Inches(0.8+ci*2.45)
            R(sl2,cx,bt2+Inches(0.35),chip_w,chip_h,D_SF,ln=D_BD)
            R(sl2,cx,bt2+Inches(0.35),chip_w,Inches(0.04),TEAL)
            tc=TEAL if not tmpl else D_DIM
            T(sl2,cx+Inches(0.15),bt2+Inches(0.5),Inches(2),Inches(0.3),ct,9.5,tc,b=True)
            T(sl2,cx+Inches(0.15),bt2+Inches(0.85),Inches(2),Inches(0.5),cd,8,D_MID if not tmpl else D_DIM,i=tmpl)

        # Band 3 — Impact
        bt3=Inches(4.9)
        if tmpl:
            for mi,(mv,md) in enumerate(mets[:3]):
                mx=Inches(0.8+mi*2.6)
                R(sl2,mx,bt3+Inches(0.3),Inches(2.3),Inches(1.0),D_SF,ln=D_BD)
                T(sl2,mx+Inches(0.2),bt3+Inches(0.4),Inches(1.2),Inches(0.4),mv,24,D_BD,b=True)
                T(sl2,mx+Inches(1.3),bt3+Inches(0.45),Inches(0.9),Inches(0.4),md,7.5,D_DIM,i=True)
            BUL(sl2,Inches(8.8),bt3+Inches(0.35),Inches(4),c3d[2:],WHITE,D_TXT,tmpl=True,tmc=D_DIM)
        else:
            for mi,(mv,md) in enumerate(mets[:3]):
                mx=Inches(0.8+mi*2.6)
                R(sl2,mx,bt3+Inches(0.3),Inches(2.3),Inches(1.0),D_SF,ln=D_BD)
                T(sl2,mx+Inches(0.2),bt3+Inches(0.4),Inches(1.2),Inches(0.4),mv,24,EMER,b=True)
                T(sl2,mx+Inches(1.3),bt3+Inches(0.45),Inches(0.9),Inches(0.4),md,7.5,D_MID)
            BUL(sl2,Inches(8.8),bt3+Inches(0.35),Inches(4),c3d[2:],WHITE,D_TXT)

        # Footer
        ft=Inches(7.1);R(sl2,Inches(0),ft,SW,Inches(0.4),D_SF)
        R(sl2,Inches(0),ft,SW,Inches(0.01),D_BD)
        R(sl2,Inches(0.6),ft+Inches(0.15),Inches(0.07),Inches(0.07),TEAL)
        T(sl2,Inches(0.8),ft+Inches(0.08),Inches(3),Inches(0.25),"Confidential \u2014 Capco",7.5,D_DIM)
        T(sl2,Inches(8),ft+Inches(0.08),Inches(5),Inches(0.25),fr,7.5,D_DIM,a=PP_ALIGN.RIGHT)

    ncontent("SDLC Transformation with BA Genie","Financial Services \u2022 Requirements Engineering \u2022 Custom AI Solution",MET,C1,C2,C3,fr="BA Genie \u2022 SDLC / Engineering")
    ncontent("[Solution Name] \u2014 [Use Case Title]","[Industry] \u2022 [Function / Process Area] \u2022 [Solution Type]",TMET,T1,T2,T3,tmpl=True,fr="[Solution Name] \u2022 [Domain]",stat_val="___",stat_lbl="Key problem\nstatistic")
    prs.save("/home/user/Intake-to-GTM/Demo One Slider - Narrative.pptx");print("Narrative done")

light()
narrative()
print("All done!")
