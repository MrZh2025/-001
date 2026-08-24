import sys
sys.stdout.reconfigure(encoding='utf-8')
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

def create_exam_word():
    doc = Document()
    
    # 页面设置 A4
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.2)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        
        # 页脚
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_run = f_p.add_run("— 第一类医疗器械实战考试卷 · 共 3 页 —")
        f_run.font.name = "宋体"
        f_run.font.size = Pt(9)
        f_run.font.color.rgb = RGBColor(128, 128, 128)

    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    doc.styles['Normal'].font.size = Pt(10.5)
    doc.styles['Normal'].font.color.rgb = RGBColor(33, 37, 41)
    
    # 试卷主标题
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run("四川省医疗器械监管与合规实战考试卷")
    run_title.font.name = "黑体"
    run_title._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    
    # 副标题
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run("科目：第一类医疗器械备案、生产与商业化全流程实操（单选+多选+判断+填空 · 满分100分）")
    run_sub.font.name = "楷体"
    run_sub._element.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
    run_sub.font.size = Pt(10.5)
    run_sub.font.bold = True

    # 考生信息栏表格
    table_info = doc.add_table(rows=1, cols=4)
    table_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_info.autofit = False
    
    col_widths = [Cm(3.8), Cm(4.2), Cm(3.8), Cm(4.2)]
    headers = ["考生姓名：________", "准考证号：________", "考核得分：________", "复核签字：________"]
    for i, cell in enumerate(table_info.rows[0].cells):
        cell.width = col_widths[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(headers[i])
        r.font.name = "宋体"
        r.font.size = Pt(10)
        r.font.bold = True
    
    def add_sec_title(title_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(title_text)
        r.font.name = "黑体"
        r.font.size = Pt(12)
        r.font.bold = True
        return p

    def add_q(q_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.25
        r = p.add_run(q_text)
        r.font.name = "宋体"
        r.font.size = Pt(10.5)
        return p

    def add_options(opt_text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.2
        r = p.add_run(opt_text)
        r.font.name = "宋体"
        r.font.size = Pt(10)
        return p

    # 一、 单选题
    add_sec_title("一、 单项选择题（共 10 题，每题 3 分，共 30 分。每题只有一个正确选项）")
    q1 = [
        ("1. 医疗器械效用的实现，主要通过以下哪种方式？（   ）",
         "A. 药理学方式        B. 免疫学方式        C. 物理等方式        D. 代谢方式"),
        ("2. 企业在办理第一类医疗器械产品备案前，唯一必须持有的法定行政资质证件是：（   ）",
         "A. 《医疗器械生产许可证》              B. 《第一类医疗器械生产备案凭证》\nC. 《医疗器械经营许可证》              D. 《企业营业执照》"),
        ("3. 依据国家药监局2022年第62号公告，在办理第一类医疗器械备案时，已取消提交、转为企业体系留存备查的资料是：（   ）",
         "A. 《产品技术要求（PTR）》             B. 《产品检验报告》\nC. 《产品风险分析资料》和《临床评价资料》 D. 《产品说明书及标签样稿》"),
        ("4. 境内第一类医疗器械产品备案的法定受理与发证部门是：（   ）",
         "A. 县/区级市场监督管理局              B. 设区的市级负责药品监督管理的部门\nC. 省级药品监督管理局                  D. 国家药品监督管理局"),
        ("5. 关于第一类医疗器械备案发证与信息公开的时间顺序，下列说法正确的是：（   ）",
         "A. 先在国家药监局公示30天，无异议后再发凭证\nB. 必须先通过专家评审会，再公示，最后拿证\nC. 先拿备案凭证，市药监部门在备案之日起5个工作日内向社会公开\nD. 备案凭证与公开没有任何时间先后关联"),
        ("6. 研发人员老周研制医用冷敷凝胶（一类）时，为了增强止痛效果，擅自添加了“薄荷脑提取液与布洛芬”，该产品将面临的合规后果是：（   ）",
         "A. 仍属于第一类医疗器械，直接备案即可\nB. 擅自添加药理成分，违规高类低报，备案将被依法取消甚至面临重大行政处罚\nC. 只要在说明书上写明含有药物，就符合一类备案要求\nD. 只要检验报告合格，即可按一类上市销售"),
        ("7. 关于第一类医疗器械在流通与销售环节的证件要求，下列说法正确的是：（   ）",
         "A. 必须办理《医疗器械经营许可证》方可销售\nB. 必须办理《第二类医疗器械经营备案凭证》方可销售\nC. 法律上根本没有“销售许可证”，且一类器械经营彻底免许可、免备案，凭营业执照即可全网销售\nD. 线上开网店销售一类器械必须取得省级药监局特许销售批文"),
        ("8. 送交第三方检测机构做全项检验的样品，其关键合规要求是：（   ）",
         "A. 必须先取得《医疗器械生产许可证》后生产的样品\nB. 必须是具有代表性的合格样品，且检验项目必须100%覆盖《产品技术要求（PTR）》所列指标\nC. 检验指标只要测外观和装量两项即可，其余可免检\nD. 只能送到国家药监局北京实验室检验，地方检验所报告无效"),
        ("9. 医疗器械委托代工生产中，品牌方（委托方）与代工厂（受托方）的核心法定责任划分是：（   ）",
         "A. 品牌方只管收钱，所有法律责任100%由代工厂承担\nB. 代工厂对产品全生命周期负责，品牌方对车间生产负责\nC. 品牌方对产品全生命周期质量安全负总责，代工厂对车间每批产品制造质量负责\nD. 双方均不对产品质量负责，由第三方检测机构负全责"),
        ("10. 只有在经营、批发、销售以下哪类产品时，才必须去市药监局办理《医疗器械经营许可证》？（   ）",
         "A. 第一类医疗器械（如冷敷凝胶、液体敷料）\nB. 第二类医疗器械（如医用口罩、额温枪、血压计）\nC. 第三类医疗器械（如输液器、注射器、心脏支架、隐形眼镜等高风险产品）\nD. 普通日用防尘口罩")
    ]
    for q, opt in q1:
        add_q(q)
        add_options(opt)

    # 二、 多选题
    add_sec_title("二、 多项选择题（共 5 题，每题 4 分，共 20 分。多选、少选、错选均不得分）")
    q2 = [
        ("1. 下列属于《第一类医疗器械产品备案申报资料》（提交给药监局的必备卷宗）的有：（   ）",
         "A. 第一类医疗器械备案表（法人签字盖章）   B. 产品技术要求（PTR）\nC. 产品全项检验报告                     D. 产品说明书和标签样稿\nE. 生产制造信息（工艺流程图等）"),
        ("2. 关于医疗器械四大核心证件的比喻与定位，下列对应正确的有：（   ）",
         "A. 产品备案证 ➔ 低风险（一类）产品的“普通身份证 / 出生证明”\nB. 生产备案证 ➔ 生产低风险（一类）产品的“工厂普通准生证”\nC. 生产许可证 ➔ 生产中高风险（二/三类）产品的“工厂高级特许准生证”\nD. 经营许可证 ➔ 倒买倒卖高风险（三类）产品的“高危特许通行证”"),
        ("3. 研发人员小王想要通过“委托代工（OEM）”模式将自己的一类产品上市，他【不需要】办理以下哪些证件？（   ）",
         "A. 企业营业执照                        B. 第一类医疗器械产品备案凭证\nC. 第一类医疗器械生产备案凭证            D. 医疗器械生产许可证\nE. 医疗器械经营许可证"),
        ("4. 依据法规，第一类医疗器械《产品技术要求（PTR）》正文必须包含的主要板块有：（   ）",
         "A. 产品型号/规格及其划分说明            B. 性能指标（客观定量、覆盖强制性标准）\nC. 检验方法（与性能指标逐一对应）        D. 附录：主要原材料及接触人体部件材质清单"),
        ("5. 当上市后的一批第一类医疗器械被药监局抽检判定为“微生物超标不合格”时，以下处理规则正确的有：（   ）",
         "A. 药监局首先向产品备案人（品牌方）下发不合格通知并启动调查\nB. 药监局会同时延伸核查受托生产的代工厂车间与检验记录\nC. 若查明是代工厂车间污染所致，代工厂需承担行政处罚并赔偿品牌方全部损失\nD. 只要找了代工厂代工，品牌方对外没有任何法律责任")
    ]
    for q, opt in q2:
        add_q(q)
        add_options(opt)

    # 三、 判断题
    add_sec_title("三、 判断分析题（共 5 题，每题 3 分，共 15 分。对的打“√”，错的打“×”）")
    q3 = [
        "1. （   ）在实验室研发打样阶段，因为还没有拿到《生产备案凭证》，所以做出来的实物样品属于“非法生产”，不能送去检测所检验。",
        "2. （   ）企业拿到《第一类医疗器械产品备案凭证》后，只要在自己的营业执照经营范围里有“一类器械销售”，在天猫、京东等网店开店卖货时，平台不会要求出示任何“经营许可证”或“销售许可证”。",
        "3. （   ）编写《产品技术要求（PTR）》时，为了图省事，性能指标栏目可以直接写“见企业随附资料”或“按供货合同标准执行”。",
        "4. （   ）无论是走委托代工（OEM）还是自建厂房自产，企业都必须先取得《第一类医疗器械产品备案凭证》，才能去推进后续的生产环节。",
        "5. （   ）医疗器械生产企业在自己的注册地址或生产地址销售本企业生产的产品，依法免予办理经营许可证或经营备案。"
    ]
    for q in q3:
        add_q(q)

    # 四、 填空题
    add_sec_title("四、 核心法规与实操填空题（共 7 题，15 个空，共 35 分）")
    q4 = [
        "1. （4分）医疗器械是指直接或间接用于人体，其效用主要通过 ____________ 等方式获得，而不是通过 ____________、免疫学或者代谢的方式获得。",
        "2. （6分）我国医疗器械法定办证先后顺序公式为：【 1. ____________ 】 ➔ 【 2. ________________________ 】 ➔ 【 3. ________________________（若自产） 】 ➔ 合法上市销售。",
        "3. （5分）第一类医疗器械产品准入实行 ____________ 性备案，法定受理与核准部门为所在地设区的 ____________ 级市场监管/药监部门。",
        "4. （6分）在委托代工（OEM）模式下，品牌方（委托方）对外对产品的 ________________ 质量安全负总责；代工厂（受托方）对车间的 ________________ 质量和每批成品出厂 ____________ 负责。",
        "5. （4分）依据《医疗器械监督管理条例》第四十一条，从事第一类医疗器械经营销售活动，不需申请 ____________ ，也无需办理 ____________ 。",
        "6. （5分）第一类医疗器械《产品技术要求（PTR）》主要包含：产品型号规格划分、性能指标、 ____________ 方法以及附录 ________________ 清单。",
        "7. （5分）老周一人研发针灸针项目，采用委托代工模式上市，个人全流程仅需办理 ______ 张法定证件，从公司注册到产品正式上市卖货，总周期预计约 ________ 个月左右。"
    ]
    for q in q4:
        p = add_q(q)
        p.paragraph_format.space_after = Pt(5)

    # 换页 - 附录答案
    doc.add_page_break()
    p_ans_title = doc.add_paragraph()
    p_ans_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_at = p_ans_title.add_run("【附录】标准参考答案（供评卷核对）")
    r_at.font.name = "黑体"
    r_at.font.size = Pt(14)
    r_at.font.bold = True

    add_sec_title("一、 单选题答案（30分）")
    doc.add_paragraph("1. C    2. D    3. C    4. B    5. C    6. B    7. C    8. B    9. C    10. C")

    add_sec_title("二、 多选题答案（20分）")
    doc.add_paragraph("1. ABCDE    2. ABCD    3. CDE    4. ABCD    5. ABC")

    add_sec_title("三、 判断题答案（15分）")
    doc.add_paragraph("1. ×    2. √    3. ×    4. √    5. √")

    add_sec_title("四、 填空题标准答案（35分）")
    doc.add_paragraph("1. 物理（或物理方式） ； 药理学（或药理）\n2. 营业执照 ； 第一类医疗器械产品备案凭证 ； 第一类医疗器械生产备案凭证\n3. 告知（或告知性） ； 市（或设区的市）\n4. 全生命周期 ； 生产制造（或生产工艺） ； 检验（或放行）\n5. 许可（或经营许可） ； 备案（或经营备案）\n6. 检验（或检验方法） ； 主要原材料（或原材料）\n7. 2（或两） ； 1.5（或1~2）")

    output_path = r"F:\2026年\梦见2026年\四川省检查员\医疗器械法规汇编知识\一类产品备案\第一类医疗器械备案与实战全流程考试卷.docx"
    doc.save(output_path)
    print("Word exam updated successfully with fill-in-the-blank format.")

if __name__ == '__main__':
    create_exam_word()
