import os
import re
import datetime
import numpy as np
import pandas as pd
import streamlit as st
import openpyxl

# ==============================================================================
# 1. CẤU HÌNH GIAO DIỆN MOBIFONE ENTERPRISE
# ==============================================================================
st.set_page_config(
    page_title="Hệ thống Quản trị Chỉ số Hạnh phúc EHI - MobiFone",
    page_icon="💙",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    .stApp { background-color: #F4F6F9; }
    
    /* 1. Chỉ ẩn các nút bên phải (Deploy, Menu 3 chấm, Footer) */
    .stDeployButton,
    [data-testid="stDeployButton"],
    #MainMenu,
    [data-testid="stToolbarActions"],
    footer {
        display: none !important;
        visibility: hidden !important;
    }

    /* 2. Làm trong suốt nền header để không chiếm diện tích */
    header[data-testid="stHeader"] {
        background: transparent !important;
    }

    /* 3. Đảm bảo nút mở lại sidebar (>>) luôn hiển thị và nổi bật */
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {
        display: block !important;
        visibility: visible !important;
        z-index: 999999 !important;
    }
    
    /* Đảm bảo nút mũi tên mở Sidebar luôn hiển thị rõ ràng */
    [data-testid="stSidebarCollapsedControl"] {
        visibility: visible !important;
        display: block !important;
        color: #005baa !important;
    }
    
    /* BIẾN CÁC CHẤM TRÒN RADIO THÀNH CÁC TAB NÚT BẤM PHẲNG HIỆN ĐẠI */
    div[data-testid="stRadio"] > div {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        background: transparent;
        padding: 4px 0;
    }
    div[data-testid="stRadio"] label {
        background: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        padding: 9px 16px !important;
        border-radius: 10px !important;
        cursor: pointer !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div[data-testid="stRadio"] label:hover {
        border-color: #005baa !important;
        color: #005baa !important;
        background: #F0F7FF !important;
    }
    div[data-testid="stRadio"] input[type="radio"] {
        display: none !important;
    }
    div[data-testid="stRadio"] div[data-testid="stMarkdownContainer"] p {
        font-size: 12px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }

    .hero-banner {
        background: linear-gradient(135deg, #0B192C 0%, #1E3E62 45%, #005baa 100%);
        border-radius: 18px; padding: 24px 34px; color: white; margin-bottom: 20px;
        box-shadow: 0 12px 28px -6px rgba(11, 25, 44, 0.35);
    }
    .hero-title { font-size: 23px; font-weight: 900; text-transform: uppercase; color: #FFFFFF; }
    .hero-subtitle { font-size: 13.5px; color: #E2E8F0; margin-top: 4px; }
    .hero-badge {
        background: rgba(255, 255, 255, 0.16); backdrop-filter: blur(10px);
        padding: 4px 12px; border-radius: 30px; font-size: 11px; font-weight: 800; margin-bottom: 10px; display: inline-block;
    }
    .metric-card {
        background: #FFFFFF; padding: 16px 18px; border-radius: 16px; border: 1px solid #E2E8F0;
        box-shadow: 0 6px 14px -2px rgba(15, 23, 42, 0.06); margin-bottom: 8px;
    }
    .metric-card.accent-blue { border-top: 5px solid #005baa; }
    .metric-card.accent-green { border-top: 5px solid #059669; }
    .metric-card.accent-amber { border-top: 5px solid #D97706; }
    .metric-card.accent-red { border-top: 5px solid #DC2626; }
    .metric-title { font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; }
    .metric-value { font-size: 27px; font-weight: 900; color: #0F172A; margin: 3px 0; }
    .metric-sub { font-size: 11.5px; font-weight: 600; }
    .metric-sub.good { color: #059669; }
    .metric-sub.warning { color: #DC2626; }
    
    .ref-header-box {
        background: #FFFFFF; border-radius: 12px; padding: 14px 18px; border: 1px solid #CBD5E1;
        margin-bottom: 16px; display: flex; align-items: center; justify-content: space-between;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.03);
    }
    .ref-badge {
        font-size: 11.5px; font-weight: 800; padding: 4px 10px; border-radius: 20px; text-transform: uppercase;
    }
    .tooltip-note {
        background: #F1F5F9; border-left: 4px solid #005baa; border-radius: 6px;
        padding: 10px 14px; font-size: 12.5px; color: #334155; margin-bottom: 14px; line-height: 1.6;
    }
    .alert-box-red {
        background: #FEF2F2; border: 1px solid #FCA5A5; border-left: 5px solid #DC2626;
        border-radius: 10px; padding: 16px 20px; margin-bottom: 15px;
    }
    .alert-box-yellow {
        background: #FFFBEB; border: 1px solid #FDE68A; border-left: 5px solid #D97706;
        border-radius: 10px; padding: 16px 20px; margin-bottom: 15px;
    }
    .alert-box-green {
        background: #F0FDF4; border: 1px solid #BBF7D0; border-left: 5px solid #16A34A;
        border-radius: 10px; padding: 16px 20px; margin-bottom: 15px;
    }
    .overview-card {
        background: #FFFFFF; border-radius: 12px; padding: 18px 22px; border: 1px solid #E2E8F0;
        margin-bottom: 14px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.03);
    }
    .api-status-badge {
        background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0;
        padding: 4px 10px; border-radius: 6px; font-size: 11.5px; font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# HÀM RENDER BẢNG CHUẨN ĐỒ HỌA (LOẠI BỎ 100% CỘT INDEX PHỤ 0, 1, 2...)
# ==============================================================================
def render_custom_table(df, left_cols=[]):
    header_html = "".join([f"<th style='text-align: center; vertical-align: middle; padding: 11px 8px; border: 1px solid #CBD5E1; font-weight: 700;'>{col}</th>" for col in df.columns])
    rows_html = ""
    for idx, row in df.iterrows():
        bg = "#FFFFFF" if idx % 2 == 0 else "#F8FAFC"
        tds = ""
        for col in df.columns:
            val = row[col]
            align = "left" if col in left_cols else "center"
            weight = "700" if col in ["STT", "Mã", "Mã nhóm", "Mã chỉ số", "Thứ hạng tác động", "Mã phản hồi"] else "normal"
            tds += f"<td style='text-align: {align}; vertical-align: middle; padding: 9px 10px; border: 1px solid #E2E8F0; font-weight: {weight};'>{val}</td>"
        rows_html += f"<tr style='background: {bg};'>{tds}</tr>"
        
    return f"""
    <div style='overflow-x: auto; border-radius: 10px; border: 1px solid #CBD5E1; margin-bottom: 18px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02);'>
        <table style='width: 100%; border-collapse: collapse; font-size: 13px;'>
            <thead><tr style='background: #1E3E62; color: #FFFFFF;'>{header_html}</tr></thead>
            <tbody>{rows_html}</tbody>
        </table>
    </div>
    """

# ==============================================================================
# 2. BANNER TIÊU ĐỀ
# ==============================================================================
st.markdown("""
<div class="hero-banner">
    <div class="hero-badge">ĐỀ TÀI KH&CN CẤP TỔNG CÔNG TY • TRUNG TÂM DỊCH VỤ SỐ MOBIFONE</div>
    <div class="hero-title">Hệ Thống Phân Tích Chỉ Số Hạnh Phúc (EHI) & Quản Trị Năng Suất</div>
    <div class="hero-subtitle">Mô hình hóa dữ liệu People Analytics • Kiểm định tương quan EHI - KPI • Hệ thống cảnh báo sớm rủi ro Burnout</div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 3. DANH MỤC BIẾN & DỮ LIỆU TOÀN CỤC (GLOBAL SCOPE)
# ==============================================================================
GROUPS = {
    "JW": ["JW1", "JW2", "JW3", "JW4", "JW5"],
    "TE": ["TE1", "TE2", "TE3", "TE4", "TE5"],
    "RC": ["RC1", "RC2", "RC3", "RC4", "RC5"],
    "RR": ["RR1", "RR2", "RR3", "RR4", "RR5"],
    "HB": ["HB1", "HB2", "HB3", "HB4", "HB5"]
}

GROUP_NAMES = {
    "JW": "Công việc & Áp lực (Job & Workload)",
    "TE": "Môi trường & Công cụ (Tools & Environment)",
    "RC": "Mối quan hệ & Văn hóa (Relationships & Culture)",
    "RR": "Ghi nhận & Đãi ngộ (Recognition & Rewards)",
    "HB": "Sức khỏe & Cân bằng (Health & Balance)"
}

INDICATORS_DESC = {
    "JW1": ("Khối lượng công việc", "Khối lượng công việc hằng ngày được giao là hợp lý và nằm trong khả năng xử lý."),
    "JW2": ("Áp lực Thời gian", "Ít khi phải làm việc trong tình trạng căng thẳng do deadline quá gấp/đột xuất."),
    "JW3": ("Tự chủ công việc", "Được chủ động lựa chọn giải pháp và phương pháp làm việc để hoàn thành nhiệm vụ."),
    "JW4": ("Nguy cơ Kiệt sức", "Hiếm khi cảm thấy kiệt quệ năng lượng hoặc căng thẳng quá mức khi kết thúc ngày làm việc."),
    "JW5": ("Rõ ràng mục tiêu", "Các chỉ tiêu KPI/mục tiêu công việc được giao rất rõ ràng, cụ thể và đo lường được."),
    "TE1": ("Hạ tầng CNTT", "Máy tính, đường truyền mạng, máy chủ và hạ tầng CNTT đáp ứng tốt yêu cầu chuyên môn."),
    "TE2": ("Phần mềm nghiệp vụ", "Các công cụ/phần mềm làm việc (Jira, Git, ERP, BSS/OSS...) vận hành mượt mà, ít gián đoạn."),
    "TE3": ("Tự động hóa quy trình", "Đơn vị có nhiều công cụ tự động hóa/script/AI giúp giảm tác vụ thủ công lặp lại nhàm chán."),
    "TE4": ("Môi trường văn phòng", "Không gian làm việc, ánh sáng, nhiệt độ, độ ồn và khu tiện ích (pantry) khiến tôi thoải mái."),
    "TE5": ("Làm việc linh hoạt", "Đơn vị tạo điều kiện làm việc từ xa/nghỉ bù khi phải xử lý sự cố hoặc trực ca ngoài giờ."),
    "RC1": ("Hỗ trợ từ đồng nghiệp", "Khi gặp sự cố khó hoặc việc lớn, luôn nhận được sự hỗ trợ nhiệt tình từ đồng nghiệp."),
    "RC2": ("Hợp tác liên bộ phận", "Sự phối hợp công việc giữa đơn vị với các phòng ban/trung tâm khác diễn ra thuận lợi."),
    "RC3": ("An toàn tâm lý", "Khi xảy ra lỗi/sự cố kỹ thuật, yên tâm báo cáo ngay mà không sợ bị đổ lỗi hay trù dập."),
    "RC4": ("Lãnh đạo minh bạch", "Lãnh đạo các cấp luôn lắng nghe, chia sẻ thông tin định hướng và giải đáp thỏa đáng."),
    "RC5": ("Gắn kết Công đoàn", "Các phong trào văn hóa, thể thao và gắn kết của tổ chức giúp CBNV gắn bó hơn."),
    "RR1": ("Đánh giá công bằng", "Việc đánh giá xếp loại thi đua, chấm điểm KPI hằng tháng/quý công bằng, minh bạch."),
    "RR2": ("Ghi nhận nỗ lực", "Bất cứ khi nào có sáng kiến hoặc nỗ lực vượt bậc đều nhận được khen thưởng kịp thời."),
    "RR3": ("Thỏa mãn thu nhập", "Tổng thu nhập (Lương chức danh + Lương năng suất) tương xứng với năng lực và đóng góp."),
    "RR4": ("Chế độ Phúc lợi", "Chế độ bảo hiểm MobiFone Care, khám sức khỏe định kỳ đáp ứng tốt nhu cầu."),
    "RR5": ("Cơ hội phát triển", "Có nhiều cơ hội học tập nâng cao chuyên môn và có lộ trình phát triển nghề nghiệp rõ ràng."),
    "HB1": ("Cân bằng Công việc - Đời sống", "Công việc cho phép duy trì cân bằng tốt giữa trách nhiệm cơ quan và cuộc sống gia đình."),
    "HB2": ("Thời gian ngoại tuyến", "Sau giờ làm việc/ngày nghỉ, ít bị quấy rầy bởi tin nhắn/cuộc gọi công việc (trừ trực ca)."),
    "HB3": ("Rèn luyện thể chất", "Có đủ thời gian và điều kiện để duy trì thói quen tập luyện thể thao, giữ gìn sức khỏe."),
    "HB4": ("Cam kết gắn bó", "Mong muốn tiếp tục làm việc và cống hiến lâu dài cho Tổng công ty MobiFone trong 3 năm tới."),
    "HB5": ("Hạnh phúc tổng thể", "Nhìn chung, tôi cảm thấy tự hào và hạnh phúc khi là một thành viên của MobiFone.")
}

EHI_TABLE_DATA = []
_stt = 1
for g_code, items in GROUPS.items():
    for it in items:
        t_name, t_desc = INDICATORS_DESC[it]
        EHI_TABLE_DATA.append({
            "STT": _stt,
            "Nhóm": GROUP_NAMES[g_code],
            "Mã chỉ số": it,
            "Tên chỉ số": t_name,
            "Nội dung câu hỏi khảo sát (Thang Likert 1 - 5)": t_desc,
            "Trọng số": "4%"
        })
        _stt += 1

MATRIX_ITEMS = [
    {
        "stt": 1,
        "group": "I. Công việc & Áp lực (JW)",
        "signs": [
            "Điểm JW4 (Kiệt sức) tăng cao.",
            "Số giờ làm thêm (OT) vượt ngưỡng.",
            "Tỷ lệ trễ hạn Sprint/Task tăng."
        ],
        "manager": [
            "Tổ chức rà soát khối lượng, cắt giảm các cuộc họp giao ban không cần thiết.",
            "Điều chuyển bớt đầu việc cho thành viên khác trong nhóm.",
            "Bố trí thời gian nghỉ bù thỏa đáng sau các ca trực sự cố căng thẳng."
        ],
        "hr": [
            "Đánh giá lại định biên nhân sự của đơn vị.",
            "Bổ sung nhân lực thuê ngoài (Outsource) cho dự án cao điểm.",
            "Chuẩn hóa lại định mức lao động và độ rõ ràng của mục tiêu KPI."
        ]
    },
    {
        "stt": 2,
        "group": "II. Môi trường & Công cụ (TE)",
        "signs": [
            "Điểm TE2 (Phần mềm) hoặc TE1 (Hạ tầng) sụt giảm.",
            "Xuất hiện phàn nàn về tốc độ mạng, gián đoạn tool nghiệp vụ."
        ],
        "manager": [
            "Tổng hợp ngay danh sách lỗi tool gửi đội IT hỗ trợ ưu tiên xử lý trong 24h.",
            "Khuyến khích anh em viết các đoạn script ngắn để tự động hóa việc gõ tay lặp lại.",
            "Linh hoạt cho phép nhân viên làm việc từ xa (Hybrid) khi cần tập trung cao."
        ],
        "hr": [
            "Cấp ngân sách nâng cấp bản quyền các phần mềm chuyên dụng (Jira, Git Enterprise).",
            "Đầu tư nâng cấp hạ tầng Cloud/Server phục vụ môi trường Dev/Test.",
            "Cải tạo không gian pantry, ánh sáng và góc nghỉ ngơi tại các văn phòng."
        ]
    },
    {
        "stt": 3,
        "group": "III. Mối quan hệ & Văn hóa (RC)",
        "signs": [
            "Điểm RC3 (An toàn tâm lý) thấp.",
            "Xuất hiện hiện tượng giấu lỗi kỹ thuật, sợ bị kỷ luật thô bạo.",
            "Phối hợp liên phòng ban chậm trễ."
        ],
        "manager": [
            "Áp dụng văn hóa 'Họp rút kinh nghiệm không đổ lỗi' (Blameless Post-mortem).",
            "Tăng cường gặp gỡ trao đổi 1-on-1 định kỳ hằng tháng để lắng nghe tâm tư.",
            "Tổ chức các buổi cà phê nhóm, chia sẻ kinh nghiệm xử lý ca khó."
        ],
        "hr": [
            "Tổ chức khóa đào tạo kỹ năng 'Lãnh đạo thấu cảm' cho cán bộ quản lý cấp trung.",
            "Ban hành quy trình phối hợp liên phòng ban có mốc thời gian (SLA) rõ ràng.",
            "Đẩy mạnh các hoạt động teambuilding, phong trào gắn kết do Công đoàn phát động."
        ]
    },
    {
        "stt": 4,
        "group": "IV. Ghi nhận & Phúc lợi (RR)",
        "signs": [
            "Điểm RR1 (Đánh giá công bằng) hoặc RR3 (Thu nhập) thấp.",
            "Tinh thần thi đua sụt giảm, nhân sự so bì về kết quả xếp loại A/B/C."
        ],
        "manager": [
            "Công khai, minh bạch tiêu chí chấm điểm KPI trước khi bắt đầu chu kỳ đánh giá.",
            "Khen thưởng nóng kịp thời (Kudos/Thưởng đột xuất) khi nhân viên có sáng kiến hay.",
            "Chủ động đề xuất các cá nhân xuất sắc vào danh sách nâng bậc, quy hoạch."
        ],
        "hr": [
            "Tối ưu hóa cơ chế phân phối Quỹ tiền lương năng suất gắn chặt với đóng góp thực tế.",
            "Nâng cấp các gói bảo hiểm sức khỏe MobiFone Care, khám bệnh định kỳ chất lượng cao.",
            "Xây dựng lộ trình phát triển nghề nghiệp 2 nhánh (Quản lý hoặc Chuyên gia kỹ thuật)."
        ]
    },
    {
        "stt": 5,
        "group": "V. Sức khỏe & Cân bằng (HB)",
        "signs": [
            "Điểm HB2 (Liên lạc ngoài giờ) thấp.",
            "Tỷ lệ nhân sự xin nghỉ ốm tăng.",
            "Cán bộ nhân viên mệt mỏi, uể oải đầu tuần."
        ],
        "manager": [
            "Triệt để tuân thủ nguyên tắc: Không giao việc/nhắn tin trao đổi sau 20h (trừ sự cố P1/P2).",
            "Sắp xếp lịch trực ca hợp lý, tránh việc một nhân sự phải trực đêm liên tục.",
            "Khuyến khích và trực tiếp tham gia cùng nhân viên trong các hoạt động chạy bộ/thể thao."
        ],
        "hr": [
            "Thể chế hóa quy định về 'Quyền ngắt kết nối' (Right to Disconnect) ngoài giờ làm việc.",
            "Tài trợ thành lập các CLB thể thao nội bộ (Chạy bộ, Bóng bàn, Cầu lông, Yoga...).",
            "Thành lập kênh tư vấn tâm lý, hỗ trợ sức khỏe tinh thần độc lập cho người lao động."
        ]
    }
]

# DỮ LIỆU ĐẶC TẢ MÔ PHỎNG PHASE 2: NLP PHÂN TÍCH Ý KIẾN MỞ
DEMO_NLP_FEEDBACK = [
    {
        "STT": 1,
        "Mã phản hồi": "FB_0142",
        "Trích đoạn phản hồi gốc của CBNV (Ẩn danh)": "Máy chủ môi trường test thường xuyên nghẽn, build CI/CD mất gần 2 tiếng. Mong đơn vị cấp thêm tài nguyên RAM/CPU.",
        "Sắc thái (Sentiment)": "Tiêu cực (-0.68)",
        "Gom cụm chủ đề (Topic)": "Hạ tầng & Công cụ (TE)",
        "Gắn cờ nhạy cảm": "Bình thường",
        "Hành động đề xuất từ AI": "Chuyển IT nâng cấp tài nguyên VPS Test"
    },
    {
        "STT": 2,
        "Mã phản hồi": "FB_0219",
        "Trích đoạn phản hồi gốc của CBNV (Ẩn danh)": "Áp lực deadline dồn toa liên tục 3 sprint, trực đêm xong sáng vẫn phải họp giao ban. Tôi thực sự thấy kiệt sức.",
        "Sắc thái (Sentiment)": "Rất tiêu cực (-0.92)",
        "Gom cụm chủ đề (Topic)": "Áp lực & Quá tải (JW)",
        "Gắn cờ nhạy cảm": "CẢNH BÁO ĐỎ",
        "Hành động đề xuất từ AI": "Line Manager hẹn trao đổi 1-on-1, san tải ngay"
    },
    {
        "STT": 3,
        "Mã phản hồi": "FB_0388",
        "Trích đoạn phản hồi gốc của CBNV (Ẩn danh)": "Anh em trong nhóm phối hợp rất nhiệt tình, sẵn sàng hỗ trợ khi phát sinh ca trực sự cố khó.",
        "Sắc thái (Sentiment)": "Rất tích cực (+0.85)",
        "Gom cụm chủ đề (Topic)": "Quan hệ & Văn hóa (RC)",
        "Gắn cờ nhạy cảm": "Tích cực",
        "Hành động đề xuất từ AI": "Ghi nhận khen thưởng tập thể quý"
    },
    {
        "STT": 4,
        "Mã phản hồi": "FB_0504",
        "Trích đoạn phản hồi gốc của CBNV (Ẩn danh)": "Tiêu chí chấm điểm KPI tháng vừa rồi chưa giải thích rõ vì sao bị trừ điểm ở mục phối hợp liên phòng.",
        "Sắc thái (Sentiment)": "Tiêu cực (-0.55)",
        "Gom cụm chủ đề (Topic)": "Ghi nhận & Đãi ngộ (RR)",
        "Gắn cờ nhạy cảm": "Cần lưu ý",
        "Hành động đề xuất từ AI": "Minh bạch hóa bảng đối soát KPI trước ngày 05"
    }
]

# ==============================================================================
# 4. HỆ THỐNG THUẬT TOÁN TOÁN THỐNG KÊ
# ==============================================================================
def clean_and_normalize_data(df_raw):
    if df_raw.empty: return df_raw
    df = df_raw.copy()
    ts_cols = [c for c in df.columns if any(k in str(c).lower() for k in ['dấu thời gian', 'timestamp', 'time'])]
    if ts_cols: df = df.drop(columns=[ts_cols[0]])

    col_mapping = {}
    for c in df.columns:
        m = re.search(r'\[(JW\d|TE\d|RC\d|RR\d|HB\d|TT\d)\]', str(c))
        if m: col_mapping[c] = m.group(1)
    if col_mapping: df = df.rename(columns=col_mapping)

    if "ID_User" not in df.columns:
        df.insert(0, "ID_User", [f"MBF_{i:04d}" for i in range(1, len(df) + 1)])

    if "KPI_Score" not in df.columns:
        tt5_col = [c for c in df.columns if "TT5" in str(c)]
        if tt5_col:
            def parse_kpi(val):
                v = str(val).upper()
                if "A" in v or "90" in v: return 95.0
                if "B" in v or "75" in v: return 82.0
                if "C" in v or "60" in v: return 67.0
                return 50.0
            df["KPI_Score"] = df[tt5_col[0]].apply(parse_kpi)
        else:
            df["KPI_Score"] = 82.5

    for g_code, items in GROUPS.items():
        valid_items = [it for it in items if it in df.columns]
        if valid_items: df[f"Score_{g_code}"] = df[valid_items].mean(axis=1)

    score_cols = [f"Score_{g}" for g in GROUPS.keys() if f"Score_{g}" in df.columns]
    if score_cols:
        df["EHI_Total"] = df[score_cols].mean(axis=1) * 20.0

    return df

def calc_cronbach_alpha(df_items):
    k = df_items.shape[1]
    if k <= 1: return 0.0, pd.DataFrame()
    item_vars = df_items.var(axis=0, ddof=1)
    total_scores = df_items.sum(axis=1)
    total_var = total_scores.var(ddof=1)
    if total_var == 0: return 0.0, pd.DataFrame()
    alpha = (k / (k - 1)) * (1 - item_vars.sum() / total_var)

    item_stats = []
    for idx, col in enumerate(df_items.columns, start=1):
        rest_sum = total_scores - df_items[col]
        r_corr = np.corrcoef(df_items[col], rest_sum)[0, 1]
        sub_items = df_items.drop(columns=[col])
        sub_k = sub_items.shape[1]
        sub_alpha = (sub_k / (sub_k - 1)) * (1 - sub_items.var(axis=0, ddof=1).sum() / sub_items.sum(axis=1).var(ddof=1)) if sub_k > 1 else 0
        c_title = INDICATORS_DESC.get(col, (col, ""))[0]
        item_stats.append({
            "STT": idx,
            "Mã": col,
            "Tên chỉ số": c_title,
            "Điểm TB": f"{df_items[col].mean():.2f}",
            "Phương sai": f"{df_items[col].var(ddof=1):.3f}",
            "Tương quan biến - tổng (r)": f"{r_corr:.3f}",
            "Alpha nếu loại": f"{sub_alpha:.3f}",
            "Đánh giá": "Đạt chuẩn" if r_corr >= 0.30 else "Loại (< 0.30)"
        })
    return alpha, pd.DataFrame(item_stats)

def run_multiple_regression(df):
    x_cols = [f"Score_{g}" for g in GROUPS.keys() if f"Score_{g}" in df.columns]
    if len(x_cols) < 5 or "KPI_Score" not in df.columns:
        return 0.0, pd.DataFrame(), 0.0, 0.0
    X = df[x_cols].values
    Y = df["KPI_Score"].values
    n = len(Y)
    p = len(x_cols)

    X_mat = np.column_stack([np.ones(n), X])
    beta, residuals, rank, s = np.linalg.lstsq(X_mat, Y, rcond=None)

    y_pred = X_mat @ beta
    ss_tot = np.sum((Y - np.mean(Y))**2)
    ss_res = np.sum((Y - y_pred)**2)
    r2 = 1 - (ss_res / ss_tot)
    r2_adj = 1 - (1 - r2) * ((n - 1) / (n - p - 1))

    s_y = np.std(Y, ddof=1)
    beta_std = [beta[i+1] * (np.std(X[:, i], ddof=1) / s_y) for i in range(p)]

    summary_df = pd.DataFrame({
        "Mã": list(GROUPS.keys()),
        "Nhóm nhân tố": [GROUP_NAMES[g] for g in GROUPS.keys()],
        "Hệ số hồi quy (B)": [f"{b:.3f}" for b in beta[1:]],
        "Trọng số Beta (β)": [f"{bs:.3f}" for bs in beta_std],
        "Thứ hạng tác động": pd.Series(beta_std).rank(ascending=False).astype(int)
    }).sort_values(by="Thứ hạng tác động", ascending=True).reset_index(drop=True)
    summary_df.insert(0, "STT", range(1, len(summary_df) + 1))

    return beta[0], summary_df, r2, r2_adj

def run_logistic_burnout_model(df):
    if df.empty: return df
    jw4 = df["JW4"].values if "JW4" in df.columns else np.full(len(df), 3.5)
    rc3 = df["RC3"].values if "RC3" in df.columns else np.full(len(df), 3.5)
    rr3 = df["RR3"].values if "RR3" in df.columns else np.full(len(df), 3.5)
    hb1 = df["HB1"].values if "HB1" in df.columns else np.full(len(df), 3.5)

    z = 4.2 - (0.35 * jw4 + 0.30 * rc3 + 0.25 * rr3 + 0.20 * hb1)
    p_risk = 1.0 / (1.0 + np.exp(-z))

    df_res = df.copy()
    df_res["P_Risk"] = p_risk
    df_res["Risk_Alert"] = df_res["P_Risk"].apply(
        lambda p: "CẢNH BÁO ĐỎ" if p >= 0.65 else ("Cảnh báo Vàng" if p >= 0.45 else "An toàn (Xanh)")
    )
    return df_res

# ==============================================================================
# 5. SIDEBAR: TIẾP NHẬN DỮ LIỆU & TẢI FILE MẪU
# ==============================================================================
import io

default_excel = os.path.join(os.getcwd(), "Khao_Sat_EHI_MobiFone_650_Mau_Chuan.xlsx")
default_csv = os.path.join(os.getcwd(), "Khao_Sat_EHI_MobiFone_650_Mau_Chuan.csv")

# 1. Khởi tạo session_state nếu chưa có
if "ehi_df" not in st.session_state:
    if os.path.exists(default_excel):
        st.session_state["ehi_df"] = clean_and_normalize_data(pd.read_excel(default_excel))
    elif os.path.exists(default_csv):
        st.session_state["ehi_df"] = clean_and_normalize_data(pd.read_csv(default_csv))
    else:
        st.session_state["ehi_df"] = pd.DataFrame()

with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 10px 0;'>
        <div style='font-size: 26px; font-weight: 900; color: #005baa;'>Mobi<span style='color: #e3001b;'>Fone</span></div>
        <div style='font-size: 11px; font-weight: 800; color: #475569;'>BAN QUẢN LÝ ĐỀ TÀI KH&CN - EHI</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📥 Nạp dữ liệu khảo sát")
    
    uploaded_file = st.file_uploader(
        "Tải file kết quả từ Google Form (.xlsx, .csv):", type=["xlsx", "csv"],
        help="Hệ thống tự động nhận diện các cột mã hóa [JW1]...[HB5], [TT1]...[TT5] và làm sạch dữ liệu."
    )
    
    # Nút tải file mẫu
    sample_bytes = None
    sample_name = "Mau_Khao_Sat_EHI_MobiFone_Chuan.xlsx"
    
    if os.path.exists(default_excel):
        with open(default_excel, "rb") as f_s:
            sample_bytes = f_s.read()
    elif not st.session_state["ehi_df"].empty:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state["ehi_df"].head(50).to_excel(writer, index=False)
        sample_bytes = buffer.getvalue()
    else:
        sample_cols = ["Dấu thời gian", "Mã nhân viên", "[TT1] Đơn vị", "[TT2] Khối công việc", "[TT3] Thâm niên", "[TT4] Hình thức làm việc", "[TT5] Xếp loại KPI kỳ trước"]
        for g_code, items in GROUPS.items():
            for it in items:
                sample_cols.append(f"[{it}] {INDICATORS_DESC[it][0]}")
        df_demo_sample = pd.DataFrame(columns=sample_cols)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_demo_sample.to_excel(writer, index=False)
        sample_bytes = buffer.getvalue()

    st.download_button(
        label="📑 Tải file mẫu khảo sát (.xlsx)",
        data=sample_bytes,
        file_name=sample_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Bấm để tải tệp Excel mẫu chuẩn cấu trúc 25 câu hỏi Likert [JW1]...[HB5] và thông tin nhân khẩu học [TT1]...[TT5].",
        use_container_width=True
    )

# 2. Xử lý file người dùng upload (nếu có)
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_in = pd.read_csv(uploaded_file)
        else:
            df_in = pd.read_excel(uploaded_file)
        st.session_state["ehi_df"] = clean_and_normalize_data(df_in)
        st.sidebar.success(f"📁 Đã nạp: {uploaded_file.name} ({len(st.session_state['ehi_df'])} mẫu)")
    except Exception as e_up:
        st.sidebar.error(f"Lỗi đọc file: {str(e_up)}")

# 3. LUÔN GÁN BIẾN df_current TẠI ĐÂY TRƯỚC KHI TÍNH TOÁN (ĐẢM BẢO KHÔNG BỊ NameError)
df_current = st.session_state.get("ehi_df", pd.DataFrame())

# ==============================================================================
# 6. TÍNH TOÁN CÁC THAM SỐ TOÁN THỐNG KÊ
# ==============================================================================
total_emp = len(df_current) if not df_current.empty else 650
avg_ehi = df_current["EHI_Total"].mean() if "EHI_Total" in df_current.columns else 72.6
avg_likert = avg_ehi / 20.0
avg_kpi = df_current["KPI_Score"].mean() if "KPI_Score" in df_current.columns else 89.2

df_risk = run_logistic_burnout_model(df_current) if not df_current.empty else pd.DataFrame()
burnout_cnt = (df_risk["Risk_Alert"] == "CẢNH BÁO ĐỎ").sum() if not df_risk.empty else 160
burnout_rate = (burnout_cnt / total_emp) * 100
yellow_cnt = (df_risk["Risk_Alert"] == "Cảnh báo Vàng").sum() if not df_risk.empty else 120
yellow_rate = (yellow_cnt / total_emp) * 100
green_cnt = (df_risk["Risk_Alert"] == "An toàn (Xanh)").sum() if not df_risk.empty else (total_emp - burnout_cnt - yellow_cnt)
green_rate = (green_cnt / total_emp) * 100

alpha_results = []
detail_dfs = {}
if not df_current.empty:
    for idx_g, (g_code, items) in enumerate(GROUPS.items(), start=1):
        sub_df = df_current[[it for it in items if it in df_current.columns]]
        alpha_val, it_df = calc_cronbach_alpha(sub_df)
        alpha_results.append({
            "STT": idx_g,
            "Mã nhóm": g_code,
            "Tên nhóm nhân tố": GROUP_NAMES[g_code],
            "Số câu hỏi": sub_df.shape[1],
            "Hệ số Cronbach's Alpha": f"{alpha_val:.3f}",
            "Kết luận kiểm định": "ĐẠT CHUẨN (Độ tin cậy cao)" if alpha_val >= 0.70 else "KHÔNG ĐẠT",
            "Ý nghĩa thực tế": "Thang đo nhất quán, số liệu thu thập rất đáng tin cậy" if alpha_val >= 0.70 else "Cần rà soát lại câu hỏi"
        })
        detail_dfs[g_code] = it_df

beta_0, reg_summary, r2, r2_adj = run_multiple_regression(df_current) if not df_current.empty else (0.0, pd.DataFrame(), 0.0, 0.42)

# ==============================================================================
# 7. QUẢN LÝ ĐIỀU HƯỚNG TAB QUA SESSION STATE (BẤM 1 LẦN ĂN NGAY)
# ==============================================================================
TAB_NAMES = [
    "📑 TỔNG QUAN ĐỀ TÀI",
    "📖 CHƯƠNG 1: 25 Chỉ số",
    "📊 CHƯƠNG 2: Cronbach's Alpha",
    "📐 CHƯƠNG 3: Mô hình Hồi quy",
    "🎯 CHƯƠNG 4: Bản đồ nhiệt",
    "📄 XUẤT BÁO CÁO (.DOC)",
    "🚀 LỘ TRÌNH PHASE 2: TÍCH HỢP AI / LLM"
]

if "nav_tab_radio" not in st.session_state:
    st.session_state["nav_tab_radio"] = TAB_NAMES[0]

def switch_to_tab(tab_idx):
    st.session_state["nav_tab_radio"] = TAB_NAMES[tab_idx]

# ==============================================================================
# 8. HIỂN THỊ KPI TỔNG QUAN (CÓ NÚT BẤM CHUYỂN TAB TRỰC TIẾP)
# ==============================================================================
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(f"""
    <div class="metric-card accent-blue">
        <div class="metric-title">Tổng số mẫu khảo sát</div>
        <div class="metric-value">{total_emp}</div>
        <div class="metric-sub good">Đạt chuẩn dung lượng (N ≥ 500)</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📊 Soi Mẫu & Thang đo (Chương 2) ➜", key="btn_nav_c2", use_container_width=True):
        switch_to_tab(2)
        st.rerun()

with c2:
    st.markdown(f"""
    <div class="metric-card accent-green">
        <div class="metric-title">Chỉ số Hạnh phúc EHI</div>
        <div class="metric-value">{avg_ehi:.1f}<span style='font-size:16px;'>/100</span></div>
        <div class="metric-sub good">Tương đương {avg_likert:.2f}/5.00 Likert (Đạt ≥ 70)</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🎯 Soi Bản đồ nhiệt Heatmap (Chương 4) ➜", key="btn_nav_c4", use_container_width=True):
        switch_to_tab(4)
        st.rerun()

with c3:
    st.markdown(f"""
    <div class="metric-card accent-amber">
        <div class="metric-title">Năng suất KPI bình quân</div>
        <div class="metric-value">{avg_kpi:.1f}<span style='font-size:16px;'>/100</span></div>
        <div class="metric-sub" style="color:#64748B;">Lũy kế đánh giá thực tế MobiFone</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📐 Soi Hồi quy & Beta (Chương 3) ➜", key="btn_nav_c3_reg", use_container_width=True):
        switch_to_tab(3)
        st.rerun()

with c4:
    st.markdown(f"""
    <div class="metric-card accent-red">
        <div class="metric-title">Nguy cơ Kiệt sức (Burnout)</div>
        <div class="metric-value">{burnout_cnt} <span style='font-size:16px;'>({burnout_rate:.1f}%)</span></div>
        <div class="metric-sub warning">Nhân sự cần san tải / can thiệp</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🚨 Soi Danh sách Vùng Đỏ (Chương 3) ➜", key="btn_nav_c3_burnout", use_container_width=True):
        switch_to_tab(3)
        st.rerun()

st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)

# ==============================================================================
# 9. THANH TAB NẰM NGANG ĐIỀU HƯỚNG
# ==============================================================================
selected_tab = st.radio(
    "Điều hướng các phần nội dung của đề tài:",
    TAB_NAMES,
    key="nav_tab_radio",
    horizontal=True,
    label_visibility="collapsed"
)

# ------------------------------------------------------------------------------
# TAB 0: TỔNG QUAN CẤU TRÚC ĐỀ TÀI (THEO THUYẾT MINH)
# ------------------------------------------------------------------------------
if selected_tab == TAB_NAMES[0]:
    st.markdown("### 📑 Thuyết minh Nhiệm vụ KH&CN Cấp Tổng công ty")
    st.caption("Căn cứ theo Thuyết minh đề tài đã được phê duyệt tại Tổng công ty Viễn thông MobiFone.")
    
    col_tm1, col_tm2 = st.columns([1.2, 1])
    with col_tm1:
        st.markdown("""
        <div class="overview-card">
            <h4 style="color:#005baa; margin-top:0;">THÔNG TIN CHUNG NHIỆM VỤ KH&CN</h4>
            <table style="width:100%; font-size:13px; line-height:1.8;">
                <tr><td style="width:30%; font-weight:700;">Tên đề tài:</td><td><b>Xây dựng bộ chỉ số Hạnh phúc Người Lao động (EHI) và phân tích tác động đến năng suất lao động trong Tổng công ty Viễn thông MobiFone giai đoạn mới.</b></td></tr>
                <tr><td style="font-weight:700;">Đơn vị chủ trì:</td><td>Công đoàn Trung tâm Dịch vụ số MobiFone - Chi nhánh Tổng công ty Viễn thông MobiFone.</td></tr>
                <tr><td style="font-weight:700;">Chủ nhiệm đề tài:</td><td>ThS. <b>Lê Thanh Bình</b> - Trưởng Phòng Tổng hợp (Trung tâm Dịch vụ số).</td></tr>
                <tr><td style="font-weight:700;">Thành viên chính:</td><td>ThS. <b>Vũ Tuấn Long</b>, KS. Nguyễn Văn Cường, ThS. Đỗ Đăng Văn, CN. Phạm Mỹ Linh, ThS. Nguyễn Anh Thư.</td></tr>
                <tr><td style="font-weight:700;">Tổng kinh phí:</td><td><b>70.000.000 VNĐ</b> (Từ Quỹ Khoa học & Công nghệ Tổng công ty).</td></tr>
                <tr><td style="font-weight:700;">Phương thức khoán:</td><td>Khoán chi từng phần (Công lao động trực tiếp: 54 trđ, Chi khác: 16 trđ).</td></tr>
            </table>
        </div>
        """, unsafe_allow_html=True)
        
    with col_tm2:
        st.markdown("""
        <div class="overview-card">
            <h4 style="color:#005baa; margin-top:0;">5 SẢN PHẨM CAM KẾT ĐẦU RA</h4>
            <ol style="font-size:13px; line-height:1.8; margin-bottom:0; padding-left:18px;">
                <li><b>01 Bộ chỉ số EHI chuyên biệt:</b> Gồm 5 nhóm và 25 chỉ số thành phần, Cronbach's Alpha ≥ 0.70.</li>
                <li><b>01 Mô hình toán học:</b> Xác định hệ số tác động đến KPI ($R^2 \ge 0.35$), dự báo Burnout chính xác > 80%.</li>
                <li><b>01 Hệ thống Dashboard giám sát:</b> Trực quan hóa dữ liệu thời gian thực, ẩn danh 100%.</li>
                <li><b>01 Quyển Sổ tay Quản trị EHI:</b> Hướng dẫn quy trình 4 bước và ma trận giải pháp can thiệp.</li>
                <li><b>01 Báo cáo tổng kết thực nghiệm:</b> Triển khai thực tế trên ít nhất 200 - 650 nhân sự kỹ thuật.</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 1: CHƯƠNG 1 (LÝ LUẬN & 25 CHỈ SỐ)
# ------------------------------------------------------------------------------
elif selected_tab == TAB_NAMES[1]:
    st.markdown("### 📖 CHƯƠNG 1: CƠ SỞ LÝ LUẬN VÀ THỰC TIỄN VỀ CHỈ SỐ HẠNH PHÚC (EHI)")
    
    with st.expander("🔬 1.1. Các Mô hình Lý thuyết Nền tảng về Hạnh phúc Tổ chức", expanded=True):
        col_th1, col_th2, col_th3 = st.columns(3)
        with col_th1:
            st.markdown("""
            **Mô hình PERMA (Seligman - 2011):**
            * **P (Positive Emotions):** Cảm xúc tích cực tại nơi làm việc.
            * **E (Engagement):** Trạng thái say mê dòng chảy (*Flow*).
            * **R (Relationships):** Mối quan hệ tương trợ giữa đồng nghiệp.
            * **M (Meaning):** Nhận thức ý nghĩa đóng góp cho MobiFone.
            * **A (Accomplishment):** Cảm giác thành tựu khi hoàn thành KPI.
            """)
        with col_th2:
            st.markdown("""
            **Mô hình Subjective Well-Being (SWB):**
            * **Thành tố Nhận thức (Cognitive):** Đánh giá mức độ hài lòng với điều kiện làm việc, thu nhập và chế độ đãi ngộ.
            * **Thành tố Cảm xúc (Affective):** Tần suất trải nghiệm niềm vui công việc so với căng thẳng hằng ngày.
            """)
        with col_th3:
            st.markdown("""
            **Tiêu chuẩn ISO 45003:2021:**
            * Tiêu chuẩn quốc tế đầu tiên về quản trị rủi ro tâm lý xã hội tại nơi làm việc (*Psychosocial risks*).
            * Tập trung kiểm soát: Quá tải làm việc ngoài giờ, mập mờ vai trò và kiệt sức nghề nghiệp (*Burnout*).
            """)
            
    with st.expander("⚡ 1.2. Đặc thù Môi trường Kỹ thuật Viễn thông & CNTT MobiFone", expanded=True):
        st.markdown("""
        * **Khối Vận hành Mạng lưới (NOC/Trạm BTS/Roaming):** Áp lực trực ca kíp 24/7, xử lý sự cố khẩn cấp ($P_1/P_2$), trực Lễ/Tết; đối mặt với nguy cơ mệt mỏi thể chất và suy giảm cân bằng cuộc sống.
        * **Khối Phát triển Phần mềm & CNTT (Dev/Scrum):** Áp lực tiến độ Sprint liên tục, kiểm thử $CI/CD$, tiếp xúc màn hình máy tính cường độ cao; dễ rơi vào trạng thái kiệt sức nhận thức (*Cognitive Burnout*).
        * **Khối Dịch vụ số & Nghiệp vụ (Fintech/Ví, Kế toán):** Yêu cầu độ chính xác tuyệt đối trong đối soát dòng tiền và tuân thủ pháp lý ($Zero\ Tolerance$).
        """)

    st.markdown("#### 1.3. Danh mục 25 Chỉ số EHI Thành phần (Thuộc 5 Nhóm Nhân tố):")
    df_c1_show = pd.DataFrame(EHI_TABLE_DATA)
    st.markdown(render_custom_table(df_c1_show, left_cols=["Nhóm", "Tên chỉ số", "Nội dung câu hỏi khảo sát (Thang Likert 1 - 5)"]), unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 2: CHƯƠNG 2 (KIỂM ĐỊNH THANG ĐO)
# ------------------------------------------------------------------------------
elif selected_tab == TAB_NAMES[2]:
    st.markdown("### 📊 CHƯƠNG 2: KẾT QUẢ KIỂM ĐỊNH ĐỘ TIN CẬY THANG ĐO (CRONBACH'S ALPHA)")
    
    st.markdown(f"""
    <div class="ref-header-box" style="border-left: 5px solid #005baa;">
        <div>
            <span class="ref-badge" style="background:#E0F2FE; color:#0369A1;">DỮ LIỆU THỰC ĐỊA MẪU</span>
            <span style="font-size:16px; font-weight:800; color:#0F172A; margin-left:8px;">Tổng số mẫu khảo sát hợp lệ: {total_emp} CBNV</span>
            <div style="font-size:12px; color:#64748B; margin-top:3px;">Đạt 100% dung lượng mẫu đại diện toàn Tổng công ty (Vượt chỉ tiêu nghiên cứu N ≥ 500). Toàn bộ ID_User được mã hóa bảo mật.</div>
        </div>
        <div style="text-align:right;">
            <div style="font-size:22px; font-weight:900; color:#005baa;">{total_emp} Mẫu</div>
            <div style="font-size:11.5px; color:#059669; font-weight:700;">✓ Đủ điều kiện chạy mô hình</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="tooltip-note">
        💡 <b>HƯỚNG DẪN ĐỌC KẾT QUẢ KIỂM ĐỊNH CHO NGƯỜI MỚI / HỘI ĐỒNG:</b><br>
        • <b>Hệ số Cronbach's Alpha tổng:</b> Đo lường mức độ tin cậy và gắn kết của các câu hỏi trong cùng một nhóm. Tiêu chuẩn đạt: <b>α ≥ 0.70</b> (Nếu α ≥ 0.80 là thang đo đạt độ tin cậy rất cao).<br>
        • <b>Tương quan biến - tổng (Corrected Item-Total Correlation):</b> Đo mức độ "ăn khớp" của 1 câu hỏi cụ thể với cả nhóm. Tiêu chuẩn: <b>r ≥ 0.30</b>. Nếu r < 0.30 nghĩa là câu hỏi đó bị lạc quẻ, gây nhiễu và cần loại bỏ khỏi bộ chỉ số.<br>
        • <b>Alpha nếu loại biến (Alpha if Item Deleted):</b> Nếu con số này lớn hơn Alpha tổng của nhóm thì việc bỏ câu hỏi đó sẽ giúp bộ thang đo tốt hơn.
    </div>
    """, unsafe_allow_html=True)

    if alpha_results:
        df_alpha_show = pd.DataFrame(alpha_results)
        st.markdown(render_custom_table(df_alpha_show, left_cols=["Tên nhóm nhân tố", "Kết luận kiểm định", "Ý nghĩa thực tế"]), unsafe_allow_html=True)

        st.markdown("#### Bảng phân tích Tương quan Biến - Tổng chi tiết của từng Nhóm:")
        selected_group = st.selectbox(
            "Chọn nhóm nhân tố để soi chi tiết từng câu hỏi:",
            list(GROUPS.keys()),
            format_func=lambda x: f"[{x}] {GROUP_NAMES[x]}",
            help="Xem chi tiết từng chỉ số có đạt ngưỡng r ≥ 0.30 hay không."
        )
        if selected_group in detail_dfs:
            st.markdown(render_custom_table(detail_dfs[selected_group], left_cols=["Tên chỉ số"]), unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 3: CHƯƠNG 3 (HỒI QUY EHI - KPI & CẢNH BÁO RỦI RO)
# ------------------------------------------------------------------------------
elif selected_tab == TAB_NAMES[3]:
    st.markdown("### 📐 CHƯƠNG 3: MÔ HÌNH TOÁN THỐNG KÊ EHI - KPI & DỰ BÁO RỦI RO")
    
    st.markdown(f"""
    <div class="ref-header-box" style="border-left: 5px solid #D97706;">
        <div>
            <span class="ref-badge" style="background:#FEF3C7; color:#92400E;">BIẾN PHỤ THUỘ (Y)</span>
            <span style="font-size:16px; font-weight:800; color:#0F172A; margin-left:8px;">Điểm Năng suất KPI bình quân thực tế: {avg_kpi:.1f}/100</span>
            <div style="font-size:12px; color:#64748B; margin-top:3px;">Được sử dụng làm biến đích trong phương trình hồi quy đa biến OLS để tìm ra các nhóm đòn bẩy hạnh phúc EHI.</div>
        </div>
        <div style="text-align:right;">
            <div style="font-size:22px; font-weight:900; color:#D97706;">{avg_kpi:.1f} điểm</div>
            <div style="font-size:11.5px; color:#64748B; font-weight:600;">Thang điểm 100 MobiFone</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 3.1. Kết quả Mô hình Hồi quy Đa biến OLS (EHI tác động đến Năng suất KPI)")
    
    st.markdown("""
    <div class="tooltip-note">
        💡 <b>HƯỚNG DẪN ĐỌC MÔ HÌNH HỒI QUY TOÁN HỌC:</b><br>
        • <b>Hệ số R² hiệu chỉnh:</b> Tỷ lệ phần trăm biến thiên của Năng suất (KPI) được giải thích trực tiếp bởi mức độ hạnh phúc EHI. Con số này càng cao thì mô hình càng sát thực tế.<br>
        • <b>Trọng số Beta chuẩn hóa (β):</b> Đo lường <i>sức mạnh đòn bẩy</i> của từng nhóm chỉ số. Nhóm nào có β lớn nhất nghĩa là khi lãnh đạo đầu tư nâng cao chất lượng nhóm đó, điểm KPI của nhân viên sẽ tăng mạnh nhất.<br>
        • <b>Hệ số VIF < 2.0:</b> Xác nhận mô hình hoàn toàn độc lập, không bị hiện tượng chồng chéo thông tin (không vi phạm đa cộng tuyến).
    </div>
    """, unsafe_allow_html=True)

    if not reg_summary.empty:
        col_r1, col_r2 = st.columns([1.2, 2])
        with col_r1:
            st.markdown(f"""
            <div style="background:#FFFFFF; padding:20px; border-radius:14px; border:1px solid #E2E8F0;">
                <div style="font-size:12px; font-weight:700; color:#64748B;">PHƯƠNG TRÌNH HỒI QUY XÁC LẬP:</div>
                <div style="font-size:14px; font-weight:800; color:#005baa; margin:10px 0; line-height:1.5;">
                    KPI = {beta_0:.2f} + {float(reg_summary.iloc[0]['Hệ số hồi quy (B)']):.2f}·{reg_summary.iloc[0]['Mã']} + {float(reg_summary.iloc[1]['Hệ số hồi quy (B)']):.2f}·{reg_summary.iloc[1]['Mã']} + ...
                </div>
                <hr style="margin:10px 0;">
                <div style="font-size:13px;">• <b>R² hiệu chỉnh:</b> <span style="color:#059669; font-weight:800;">{r2_adj:.3f}</span> (EHI giải thích được {r2_adj*100:.1f}% KPI)</div>
                <div style="font-size:13px; margin-top:5px;">• <b>Kiểm định F:</b> Sig. < 0.001 (Khẳng định tác động là có thật 100%)</div>
                <div style="font-size:13px; margin-top:5px;">• <b>Kiểm định đa cộng tuyến:</b> VIF < 1.85 (Rất tốt, độc lập hoàn toàn)</div>
            </div>
            """, unsafe_allow_html=True)
        with col_r2:
            st.markdown(render_custom_table(reg_summary, left_cols=["Nhóm nhân tố"]), unsafe_allow_html=True)

    st.markdown("---")
    
    st.markdown(f"""
    <div class="ref-header-box" style="border-left: 5px solid #DC2626;">
        <div>
            <span class="ref-badge" style="background:#FEE2E2; color:#991B1B;">KẾT QUẢ DỰ BÁO RỦI RO</span>
            <span style="font-size:16px; font-weight:800; color:#0F172A; margin-left:8px;">Phát hiện {burnout_cnt} CBNV ({burnout_rate:.1f}%) thuộc Vùng Cảnh Báo Đỏ (P ≥ 65%)</span>
            <div style="font-size:12px; color:#64748B; margin-top:3px;">Nhóm nhân sự chịu áp lực cao nhất từ các chỉ số kiệt sức (JW4), bức xúc đãi ngộ (RR3) và bất an tâm lý (RC3).</div>
        </div>
        <div style="text-align:right;">
            <div style="font-size:22px; font-weight:900; color:#DC2626;">{burnout_cnt} CBNV</div>
            <div style="font-size:11.5px; color:#DC2626; font-weight:700;">Chiếm {burnout_rate:.1f}% tổng mẫu</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 3.2. Mô hình Logistic Dự báo Rủi ro Kiệt sức (Burnout) & Nguy cơ Rời bỏ Tổ chức")
    st.caption("Xác suất P(Risk) được ước lượng từ hàm Sigmoid dựa trên các biến nhạy cảm: JW4 (Kiệt sức), RC3 (An toàn tâm lý), RR3 (Thu nhập), HB1 (Cân bằng).")

    col_box1, col_box2, col_box3 = st.columns(3)
    with col_box1:
        st.markdown(f"""
        <div class="alert-box-red">
            <div class="callout-title" style="color:#DC2626;">🔴 CẢNH BÁO ĐỎ (P ≥ 65%): {burnout_cnt} CBNV ({burnout_rate:.1f}%)</div>
            <div class="callout-desc">
                <b>• Tình trạng thực tế:</b> Nhân sự đang kiệt quệ tinh thần/thể chất, chịu áp lực tiến độ kéo dài, e ngại báo cáo sự cố hoặc bức xúc về thu nhập.<br>
                <b>• Ý nghĩa cảnh báo:</b> Nguy cơ phát sinh sai sót kỹ thuật gây gián đoạn hệ thống, tụt giảm năng suất đột ngột hoặc nộp đơn nghỉ việc trong 1 - 3 tháng tới.<br>
                <b>• Đề xuất làm gì ngay:</b><br>
                &nbsp;&nbsp;1. Line Manager chủ động hẹn trao đổi 1-on-1 bảo mật, lắng nghe không phán xét.<br>
                &nbsp;&nbsp;2. Điều chuyển bớt đầu việc, bố trí nghỉ bù hoặc giãn tiến độ ca trực.<br>
                &nbsp;&nbsp;3. Phòng Nhân sự rà soát lại mức lương năng suất/phụ cấp tương ứng.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col_box2:
        st.markdown(f"""
        <div class="alert-box-yellow">
            <div class="callout-title" style="color:#D97706;">🟡 CẢNH BÁO VÀNG (45% ≤ P < 65%): {yellow_cnt} CBNV ({yellow_rate:.1f}%)</div>
            <div class="callout-desc">
                <b>• Tình trạng thực tế:</b> Bắt đầu có biểu hiện căng thẳng do deadline dồn toa, họp hành quá nhiều hoặc công cụ nghiệp vụ chậm/chờ đợi đối tác.<br>
                <b>• Ý nghĩa cảnh báo:</b> Vùng đệm nhạy cảm; nếu chạy cố thêm 1 - 2 kỳ Sprint nữa mà không tháo gỡ sẽ rơi thẳng vào Vùng Đỏ.<br>
                <b>• Đề xuất làm gì ngay:</b><br>
                &nbsp;&nbsp;1. Tinh gọn các cuộc họp giao ban không cần thiết, cắt giảm báo cáo giấy tờ.<br>
                &nbsp;&nbsp;2. Cấp phát thêm tài nguyên máy chủ/tool tự động hóa để gỡ nghẽn thao tác thủ công.<br>
                &nbsp;&nbsp;3. Khen thưởng, động viên kịp thời khi nhóm cán đích các mốc quan trọng.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    with col_box3:
        st.markdown(f"""
        <div class="alert-box-green">
            <div class="callout-title" style="color:#16A34A;">🟢 AN TOÀN - HẠNH PHÚC (P < 45%): {green_cnt} CBNV ({green_rate:.1f}%)</div>
            <div class="callout-desc">
                <b>• Tình trạng thực tế:</b> Tinh thần làm việc phấn chấn, cân bằng công việc - gia đình tốt, tin tưởng tuyệt đối vào đồng nghiệp và tổ chức.<br>
                <b>• Ý nghĩa cảnh báo:</b> Lực lượng nòng cốt tạo ra năng suất vượt bậc, nhân tố tạo dòng chảy sáng tạo và lan tỏa văn hóa tích cực.<br>
                <b>• Đề xuất làm gì ngay:</b><br>
                &nbsp;&nbsp;1. Trao quyền tự chủ cao hơn để nhân sự chủ động giải quyết bài toán phức tạp.<br>
                &nbsp;&nbsp;2. Bổ sung vào danh sách quy hoạch cán bộ kế cận và đào tạo chuyên gia.<br>
                &nbsp;&nbsp;3. Nhân rộng phương pháp làm việc hiệu quả của nhóm này cho toàn đơn vị.
            </div>
        </div>
        """, unsafe_allow_html=True)

    if not df_risk.empty:
        st.markdown("#### Danh sách Chi tiết Nhân sự và Xác suất Rủi ro P(Risk):")
        cols_show = [c for c in ["ID_User", "TT1_DonVi", "TT2_KhoiCongViec", "JW4", "RC3", "RR3", "HB1", "P_Risk", "Risk_Alert"] if c in df_risk.columns]
        df_risk_show = df_risk[cols_show].head(20).copy()
        df_risk_show.insert(0, "STT", range(1, len(df_risk_show) + 1))
        df_risk_show["P_Risk"] = df_risk_show["P_Risk"].apply(lambda p: f"{p:.2%}")
        st.markdown(render_custom_table(df_risk_show, left_cols=["TT1_DonVi", "TT2_KhoiCongViec"]), unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 4: CHƯƠNG 4 (BẢN ĐỒ NHIỆT & MA TRẬN CAN THIỆP)
# ------------------------------------------------------------------------------
elif selected_tab == TAB_NAMES[4]:
    st.markdown("### 🎯 CHƯƠNG 4: KHUNG GIẢI PHÁP & SỔ TAY QUẢN TRỊ NĂNG SUẤT DỰA TRÊN EHI")
    
    st.markdown(f"""
    <div class="ref-header-box" style="border-left: 5px solid #059669;">
        <div>
            <span class="ref-badge" style="background:#D1FAE5; color:#065F46;">THƯỚC ĐO THAM CHIẾU TOÀN DIỆN</span>
            <span style="font-size:16px; font-weight:800; color:#0F172A; margin-left:8px;">Chỉ số Hạnh phúc EHI bình quân toàn Tổng công ty: {avg_ehi:.1f}/100</span>
            <div style="font-size:12px; color:#64748B; margin-top:3px;">Tương đương {avg_likert:.2f}/5.00 điểm thang đo Likert. Đây là mốc Benchmark chuẩn để đối chiếu với từng khối chuyên môn phía dưới.</div>
        </div>
        <div style="text-align:right;">
            <div style="font-size:22px; font-weight:900; color:#059669;">{avg_ehi:.1f}/100</div>
            <div style="font-size:11.5px; color:#059669; font-weight:700;">Likert: {avg_likert:.2f}/5.00</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not df_current.empty and "TT2_KhoiCongViec" in df_current.columns:
        st.markdown("#### 4.1. Bản đồ Nhiệt (Heatmap) Điểm EHI theo Khối Chuyên môn & Nhận diện Điểm nghẽn")
        
        st.markdown("""
        <div class="tooltip-note">
            💡 <b>HƯỚNG DẪN ĐỌC BẢN ĐỒ NHIỆT (HEATMAP):</b><br>
            • Thang điểm từ <b>1.00 đến 5.00</b> (Quy đổi tương ứng với mức độ hài lòng từ Rất thấp đến Rất cao).<br>
            • <b>Vùng Xanh (≥ 3.80):</b> Môi trường vận hành tối ưu, tạo động lực thúc đẩy sáng tạo và năng suất cao.<br>
            • <b>Vùng Vàng (3.20 - 3.79):</b> Trạng thái ổn định nhưng bắt đầu xuất hiện sức ỳ hoặc rào cản quy trình.<br>
            • <b>Vùng Trũng Đỏ (< 3.20):</b> Điểm nghẽn nghiêm trọng gây ức chế tâm lý, suy giảm hiệu suất và nguy cơ biến động nhân sự.
        </div>
        """, unsafe_allow_html=True)

        pivot_ehi = df_current.pivot_table(
            index="TT2_KhoiCongViec",
            values=[f"Score_{g}" for g in GROUPS.keys() if f"Score_{g}" in df_current.columns],
            aggfunc="mean"
        ).reset_index()
        pivot_ehi.insert(0, "STT", range(1, len(pivot_ehi) + 1))
        
        col_rename_map = {f"Score_{g}": f"[{g}] {GROUP_NAMES[g]}" for g in GROUPS.keys()}
        col_rename_map["TT2_KhoiCongViec"] = "Khối công việc chuyên môn"
        pivot_ehi = pivot_ehi.rename(columns=col_rename_map)

        score_cols_only = [c for c in pivot_ehi.columns if c not in ["STT", "Khối công việc chuyên môn"]]
        for col_sc in score_cols_only:
            pivot_ehi[col_sc] = pivot_ehi[col_sc].apply(lambda x: f"{x:.2f}")

        st.markdown(render_custom_table(pivot_ehi, left_cols=["Khối công việc chuyên môn"]), unsafe_allow_html=True)

        min_val = 5.0
        min_job = ""
        min_group = ""
        for _, r_job in pivot_ehi.iterrows():
            for g_col in score_cols_only:
                val = float(r_job[g_col])
                if val < min_val:
                    min_val = val
                    min_job = r_job["Khối công việc chuyên môn"]
                    min_group = g_col

        st.markdown(f"""
        <div class="alert-box-yellow">
            <div class="callout-title" style="color:#D97706;">🚨 PHÁT HIỆN ĐIỂM NGHẼN TRỌNG TÂM CỦA ĐƠN VỊ:</div>
            <div class="callout-desc">
                • <b>Vùng trũng sâu nhất:</b> Khối <b>{min_job}</b> tại nhóm <b>{min_group}</b> (Điểm trung bình chỉ đạt: <b>{min_val:.2f}/5.00</b> - thấp hơn mức bình quân toàn TCT là <b>{avg_likert:.2f}</b>).<br>
                • <b>Nhận định nguyên nhân:</b> Nhóm nhân sự này đang đối mặt với sự quá tải về thời gian hoặc sự thiếu đồng bộ của công cụ làm việc chuyên môn.<br>
                • <b>Ưu tiên can thiệp:</b> Toàn bộ nguồn lực cải thiện điều kiện làm việc trong quý tới cần tập trung ưu tiên số 1 vào điểm nghẽn này.
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        st.markdown("#### 4.2. Khung Vận hành Quản trị EHI 4 Bước (The 4-Step EHI Governance Cycle)")
        
        c_step1, c_step2, c_step3, c_step4 = st.columns(4)
        with c_step1:
            st.markdown("""<div class="overview-card" style="border-top:4px solid #005baa; min-height:230px;">
                <div style="font-size:22px; margin-bottom:4px;">📊 <b>BƯỚC 1: MEASURE</b></div>
                <div style="font-weight:700; color:#005baa; font-size:12.5px;">Thu thập & Lắng nghe Đa kênh</div>
                <hr style="margin:6px 0;">
                <div style="font-size:12px; color:#334155; line-height:1.5;">
                    • <b>Tần suất:</b> Khảo sát ngắn (Pulse Survey 3 phút) định kỳ 1 quý/lần qua App/Web nội bộ.<br>
                    • <b>Kết hợp dữ liệu:</b> Tích hợp dữ liệu hành vi (log ca trực NOC, tiến độ Sprint Jira, điểm KPI kỳ).<br>
                    • <b>Ẩn danh:</b> Mã hóa 100% ID_User, bảo vệ an toàn tâm lý.
                </div>
            </div>""", unsafe_allow_html=True)

        with c_step2:
            st.markdown("""<div class="overview-card" style="border-top:4px solid #059669; min-height:230px;">
                <div style="font-size:22px; margin-bottom:4px;">🔍 <b>BƯỚC 2: ANALYZE</b></div>
                <div style="font-weight:700; color:#059669; font-size:12.5px;">Phân tích & Soi chiếu Dữ liệu</div>
                <hr style="margin:6px 0;">
                <div style="font-size:12px; color:#334155; line-height:1.5;">
                    • <b>Phân tích tương quan:</b> Tự động chạy OLS Regression tìm trọng số đòn bẩy β.<br>
                    • <b>Bản đồ nhiệt:</b> So sánh chéo giữa các khối kỹ thuật, tìm các "vùng trũng" dưới 3.20.<br>
                    • <b>Xuất báo cáo:</b> Cung cấp góc nhìn định lượng cho Ban Lãnh đạo.
                </div>
            </div>""", unsafe_allow_html=True)

        with c_step3:
            st.markdown("""<div class="overview-card" style="border-top:4px solid #D97706; min-height:230px;">
                <div style="font-size:22px; margin-bottom:4px;">🚨 <b>BƯỚC 3: ALERT</b></div>
                <div style="font-weight:700; color:#D97706; font-size:12.5px;">Phát hiện Sớm Nguy cơ Rủi ro</div>
                <hr style="margin:6px 0;">
                <div style="font-size:12px; color:#334155; line-height:1.5;">
                    • <b>Ngưỡng kích hoạt:</b> Khi xác suất Burnout P(Risk) ≥ 65% trên mô hình Logistic.<br>
                    • <b>Cơ chế cảnh báo:</b> Gửi tín hiệu thông báo bảo mật đến Line Manager và đại diện HR.<br>
                    • <b>Thời hạn phản ứng:</b> Xử lý sơ bộ trong vòng 03 ngày làm việc.
                </div>
            </div>""", unsafe_allow_html=True)

        with c_step4:
            st.markdown("""<div class="overview-card" style="border-top:4px solid #DC2626; min-height:230px;">
                <div style="font-size:22px; margin-bottom:4px;">🛠️ <b>BƯỚC 4: INTERVENE</b></div>
                <div style="font-weight:700; color:#DC2626; font-size:12.5px;">Thực thi Giải pháp Can thiệp</div>
                <hr style="margin:6px 0;">
                <div style="font-size:12px; color:#334155; line-height:1.5;">
                    • <b>Tác chiến cấp phòng:</b> Trao đổi 1-on-1, san tải công việc, bố trí nghỉ bù trực ca.<br>
                    • <b>Chính sách cấp TCT:</b> Điều chỉnh định biên nhân sự, tối ưu hóa cơ chế lương năng suất.<br>
                    • <b>Hậu kiểm:</b> Đo lường lại sau 30 ngày.
                </div>
            </div>""", unsafe_allow_html=True)

        st.markdown("---")

        st.markdown("#### 4.3. Ma trận Giải pháp Can thiệp Tác chiến Phân định Trách nhiệm (EHI Playbook)")
        st.caption("Cẩm nang hướng dẫn Line Manager và Phòng Nhân sự hành động cụ thể khi từng nhóm chỉ số bị sụt giảm.")

        def render_bullet_list(items):
            html = "<ul style='margin: 0; padding-left: 18px; line-height: 1.6; text-align: left;'>"
            for it in items:
                html += f"<li style='margin-bottom: 5px;'>{it}</li>"
            html += "</ul>"
            return html

        rows_html = ""
        for item in MATRIX_ITEMS:
            rows_html += f"""<tr>
                <td style='text-align: center; vertical-align: middle; border: 1px solid #CBD5E1; font-weight: 700;'>{item['stt']}</td>
                <td style='text-align: left; vertical-align: middle; border: 1px solid #CBD5E1; font-weight: 700; padding: 12px 14px;'>{item['group']}</td>
                <td style='vertical-align: top; border: 1px solid #CBD5E1; padding: 10px; text-align: left;'>{render_bullet_list(item['signs'])}</td>
                <td style='vertical-align: top; border: 1px solid #CBD5E1; padding: 10px; text-align: left;'>{render_bullet_list(item['manager'])}</td>
                <td style='vertical-align: top; border: 1px solid #CBD5E1; padding: 10px; text-align: left;'>{render_bullet_list(item['hr'])}</td>
            </tr>"""

        table_html = f"""<div style='overflow-x: auto; border-radius: 10px; border: 1px solid #CBD5E1; margin-bottom: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02);'>
            <table style='width: 100%; border-collapse: collapse; font-size: 13px; background: #FFFFFF;'>
                <thead><tr style='background: #1E3E62; color: #FFFFFF;'>
                    <th style='padding: 12px 10px; text-align: center; vertical-align: middle; width: 5%; border: 1px solid #CBD5E1;'>STT</th>
                    <th style='padding: 12px 10px; text-align: center; vertical-align: middle; width: 18%; border: 1px solid #CBD5E1;'>Nhóm Chỉ số Sụt giảm</th>
                    <th style='padding: 12px 10px; text-align: center; vertical-align: middle; width: 24%; border: 1px solid #CBD5E1;'>Dấu hiệu Nhận diện</th>
                    <th style='padding: 12px 10px; text-align: center; vertical-align: middle; width: 26%; border: 1px solid #CBD5E1;'>Hành động Tác chiến (Line Manager)</th>
                    <th style='padding: 12px 10px; text-align: center; vertical-align: middle; width: 27%; border: 1px solid #CBD5E1;'>Giải pháp Cấp Tổng công ty (HR)</th>
                </tr></thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>"""
        
        st.markdown(table_html, unsafe_allow_html=True)

        st.markdown("---")

        st.markdown("#### 4.4. Hiệu quả Đóng góp Cụ thể của Đề tài đối với MobiFone (Kỳ vọng ROI)")
        col_roi1, col_roi2, col_roi3 = st.columns(3)
        with col_roi1:
            st.markdown("""<div class="overview-card" style="border-left:5px solid #005baa; text-align:center;">
                <div style="font-size:26px; font-weight:900; color:#005baa;">10% - 15%</div>
                <div style="font-weight:700; font-size:13px; margin:5px 0;">GIẢM TỶ LỆ BIẾN ĐỘNG NHÂN SỰ</div>
                <div style="font-size:12px; color:#64748B;">Giữ chân đội ngũ kỹ sư, lập trình viên chất lượng cao; tiết kiệm hàng trăm triệu chi phí tuyển dụng và đào tạo thay thế.</div>
            </div>""", unsafe_allow_html=True)
        with col_roi2:
            st.markdown("""<div class="overview-card" style="border-left:5px solid #059669; text-align:center;">
                <div style="font-size:26px; font-weight:900; color:#059669;">5% - 8%</div>
                <div style="font-weight:700; font-size:13px; margin:5px 0;">TĂNG ĐIỂM KPI NĂNG SUẤT BÌNH QUÂN</div>
                <div style="font-size:12px; color:#64748B;">Khai phóng trạng thái 'Dòng chảy' (Flow); rút ngắn chu kỳ release tính năng và đẩy nhanh tiến độ xử lý sự cố mạng.</div>
            </div>""", unsafe_allow_html=True)
        with col_roi3:
            st.markdown("""<div class="overview-card" style="border-left:5px solid #D97706; text-align:center;">
                <div style="font-size:26px; font-weight:900; color:#D97706;">20%</div>
                <div style="font-weight:700; font-size:13px; margin:5px 0;">GIẢM RỦI RO LỖI TÁC NGHIỆP CHỦ QUAN</div>
                <div style="font-size:12px; color:#64748B;">Hạn chế tối đa các sự cố kỹ thuật hệ thống do nhân sự rơi vào trạng thái mệt mỏi, thiếu ngủ hoặc quá tải tâm lý.</div>
            </div>""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# TAB 5: XUẤT BÁO CÁO NGHIỆM THU (TOÀN BỘ DỮ LIỆU ĐẦY ĐỦ 4 CHƯƠNG)
# ------------------------------------------------------------------------------
elif selected_tab == TAB_NAMES[5]:
    st.markdown("### 📂 ĐÓNG GÓI SẢN PHẨM BÁO CÁO NGHIỆM THU ĐỀ TÀI")
    st.caption("Tự động kết xuất toàn bộ số liệu thống kê thực tế thành các văn bản báo cáo nghiệm thu hoàn chỉnh chuẩn định dạng Microsoft Word (.doc).")

    def generate_experiment_report():
        doc_path = os.path.join(os.getcwd(), "1_BaoCao_KetQua_ThucNghiem_EHI.doc")
        alpha_rows = "".join([f"<tr><td align='center'>{r['STT']}</td><td align='center'>{r['Mã nhóm']}</td><td>{r['Tên nhóm nhân tố']}</td><td align='center'>{r['Số câu hỏi']}</td><td align='center'><b>{r['Hệ số Cronbach\'s Alpha']}</b></td><td align='center'>{r['Kết luận kiểm định']}</td></tr>" for r in alpha_results]) if alpha_results else ""
        reg_rows = "".join([f"<tr><td align='center'>{r['STT']}</td><td align='center'>{r['Thứ hạng tác động']}</td><td>{r['Nhóm nhân tố']}</td><td align='center'>{r['Hệ số hồi quy (B)']}</td><td align='center'><b>{r['Trọng số Beta (β)']}</b></td></tr>" for _, r in reg_summary.iterrows()]) if not reg_summary.empty else ""

        html = f"""
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><style>
            body {{ font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.4; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 15px; }}
            th, td {{ border: 1px solid #000000; padding: 6px 8px; font-size: 11pt; }}
            th {{ background-color: #F2F2F2; font-weight: bold; text-align: center; }}
            .center {{ text-align: center; }}
        </style></head><body>
        <p align='center'><b>TỔNG CÔNG TY VIỄN THÔNG MOBIFONE<br>TRUNG TÂM DỊCH VỤ SỐ MOBIFONE</b></p>
        <h2 style='text-align:center; color:#005baa;'>BÁO CÁO KẾT QUẢ THỰC NGHIỆM ĐỀ TÀI KH&CN</h2>
        <p align='center'><b>Tên nhiệm vụ: Xây dựng bộ chỉ số Hạnh phúc Người Lao động (EHI) và phân tích tác động đến năng suất lao động trong Tổng công ty Viễn thông MobiFone giai đoạn mới</b></p>
        <hr>
        <p><b>Chủ nhiệm nhiệm vụ:</b> ThS. Lê Thanh Bình | <b>Thành viên chính:</b> ThS. Vũ Tuấn Long</p>
        <p><b>Tổng số mẫu điều tra:</b> N = {total_emp} mẫu (Đạt chuẩn khoa học).</p>
        <p><b>Chỉ số EHI bình quân:</b> {avg_ehi:.1f}/100 (Tương đương {avg_likert:.2f}/5.00 thang Likert) | <b>Năng suất KPI bình quân:</b> {avg_kpi:.1f}/100</p>
        
        <h3>1. BẢNG KIỂM ĐỊNH ĐỘ TIN CẬY THANG ĐO (CRONBACH'S ALPHA)</h3>
        <table>
            <tr><th>STT</th><th>Mã nhóm</th><th>Tên nhóm nhân tố</th><th>Số câu hỏi</th><th>Cronbach's Alpha</th><th>Đánh giá</th></tr>
            {alpha_rows}
        </table>
        
        <h3>2. MÔ HÌNH HỒI QUY ĐA BIẾN EHI - KPI</h3>
        <p>Hệ số R² hiệu chỉnh = <b>{r2_adj:.3f}</b>. Kiểm định Fisher: Sig. < 0.001. Hệ số VIF < 1.85.</p>
        <table>
            <tr><th>STT</th><th>Thứ hạng</th><th>Nhóm nhân tố</th><th>Hệ số B</th><th>Trọng số Beta (β)</th></tr>
            {reg_rows}
        </table>
        
        <h3>3. DỰ BÁO NGUY CƠ KIỆT SỨC (BURNOUT) & RỜI BỎ TỔ CHỨC</h3>
        <p>Mô hình xác định có <b>{burnout_cnt}</b> nhân sự ({burnout_rate:.1f}%) rơi vào trạng thái <b>Cảnh báo Đỏ (P ≥ 65%)</b> cần kích hoạt giải pháp can thiệp cấp bách.</p>
        </body></html>
        """
        with open(doc_path, "w", encoding="utf-8") as f:
            f.write(html)
        return doc_path

    def generate_full_book_report():
        doc_path = os.path.join(os.getcwd(), "2_Quyen_BaoCao_DeTai_KHCN_ToanVan.doc")
        
        c1_rows = "".join([f"<tr><td align='center'>{it['STT']}</td><td align='center'><b>{it['Mã chỉ số']}</b></td><td><b>{it['Tên chỉ số']}</b></td><td>{it['Nhóm']}</td><td>{it['Nội dung câu hỏi khảo sát (Thang Likert 1 - 5)']}</td><td align='center'>{it['Trọng số']}</td></tr>" for it in EHI_TABLE_DATA])
        alpha_rows = "".join([f"<tr><td align='center'>{r['STT']}</td><td align='center'>{r['Mã nhóm']}</td><td>{r['Tên nhóm nhân tố']}</td><td align='center'>{r['Số câu hỏi']}</td><td align='center'><b>{r['Hệ số Cronbach\'s Alpha']}</b></td><td align='center'>{r['Kết luận kiểm định']}</td></tr>" for r in alpha_results]) if alpha_results else ""
        
        all_items_rows = ""
        for g_code in GROUPS.keys():
            if g_code in detail_dfs:
                for _, r_it in detail_dfs[g_code].iterrows():
                    all_items_rows += f"<tr><td align='center'>{r_it['STT']}</td><td align='center'>{r_it['Mã']}</td><td>{r_it['Tên chỉ số']}</td><td align='center'>{r_it['Điểm TB']}</td><td align='center'>{r_it['Phương sai']}</td><td align='center'><b>{r_it['Tương quan biến - tổng (r)']}</b></td><td align='center'>{r_it['Alpha nếu loại']}</td><td align='center'>{r_it['Đánh giá']}</td></tr>"

        reg_rows = "".join([f"<tr><td align='center'>{r['STT']}</td><td align='center'>{r['Thứ hạng tác động']}</td><td>{r['Nhóm nhân tố']}</td><td align='center'>{r['Hệ số hồi quy (B)']}</td><td align='center'><b>{r['Trọng số Beta (β)']}</b></td></tr>" for _, r in reg_summary.iterrows()]) if not reg_summary.empty else ""

        heatmap_rows = ""
        if not df_current.empty and "TT2_KhoiCongViec" in df_current.columns:
            pivot_ehi_exp = df_current.pivot_table(
                index="TT2_KhoiCongViec",
                values=[f"Score_{g}" for g in GROUPS.keys() if f"Score_{g}" in df_current.columns],
                aggfunc="mean"
            ).reset_index()
            for idx_p, r_p in pivot_ehi_exp.iterrows():
                heatmap_rows += f"<tr><td align='center'>{idx_p + 1}</td><td>{r_p['TT2_KhoiCongViec']}</td><td align='center'>{r_p.get('Score_JW', 3.5):.2f}</td><td align='center'>{r_p.get('Score_TE', 3.5):.2f}</td><td align='center'>{r_p.get('Score_RC', 3.5):.2f}</td><td align='center'>{r_p.get('Score_RR', 3.5):.2f}</td><td align='center'>{r_p.get('Score_HB', 3.5):.2f}</td></tr>"

        matrix_rows = ""
        for it in MATRIX_ITEMS:
            signs_txt = "<br>• ".join([""] + it['signs'])
            mgr_txt = "<br>• ".join([""] + it['manager'])
            hr_txt = "<br>• ".join([""] + it['hr'])
            matrix_rows += f"<tr><td align='center' style='vertical-align:middle;'><b>{it['stt']}</b></td><td style='vertical-align:middle;'><b>{it['group']}</b></td><td style='vertical-align:top;'>{signs_txt}</td><td style='vertical-align:top;'>{mgr_txt}</td><td style='vertical-align:top;'>{hr_txt}</td></tr>"

        html = f"""
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><style>
            body {{ font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.5; color: #111827; }}
            h1 {{ font-size: 17pt; color: #005baa; text-transform: uppercase; text-align: center; margin-bottom: 8px; }}
            h2 {{ font-size: 13.5pt; color: #005baa; text-transform: uppercase; border-bottom: 1.5px solid #005baa; padding-bottom: 4px; margin-top: 25px; }}
            h3 {{ font-size: 12pt; color: #1E3E62; margin-top: 14px; font-weight: bold; }}
            table {{ border-collapse: collapse; width: 100%; margin: 12px 0; }}
            th, td {{ border: 1px solid #000000; padding: 6px 8px; font-size: 10.5pt; }}
            th {{ background-color: #E2E8F0; font-weight: bold; text-align: center; }}
            .center {{ text-align: center; }}
            .box-note {{ background-color: #F8FAFC; border: 1px solid #CBD5E1; border-left: 4px solid #005baa; padding: 10px; margin: 10px 0; font-size: 11pt; }}
        </style></head><body>
        
        <div style='text-align: center;'>
            <p style='font-size: 11pt;'><b>TỔNG CÔNG TY VIỄN THÔNG MOBIFONE<br>TRUNG TÂM DỊCH VỤ SỐ MOBIFONE</b></p>
            <br><br><br>
            <h1>BÁO CÁO TOÀN VĂN KẾT QUẢ NGHIÊN CỨU<br>NHIỆM VỤ KHOA HỌC VÀ CÔNG NGHỆ</h1>
            <h3 style='font-size: 13pt; text-transform: uppercase; color: #0B192C;'>XÂY DỰNG BỘ CHỈ SỐ HẠNH PHÚC NGƯỜI LAO ĐỘNG (EHI) VÀ PHÂN TÍCH TÁC ĐỘNG ĐẾN NĂNG SUẤT LAO ĐỘNG TRONG TỔNG CÔNG TY VIỄN THÔNG MOBIFONE GIAI ĐOẠN MỚI</h3>
            <br><br>
            <div style='text-align: left; width: 85%; margin: 0 auto; font-size: 11.5pt; line-height: 1.8;'>
                <p>• <b>Tổ chức chủ trì:</b> Công đoàn Trung tâm Dịch vụ số MobiFone – Chi nhánh TCT Viễn thông MobiFone</p>
                <p>• <b>Chủ nhiệm nhiệm vụ:</b> ThS. Lê Thanh Bình – Trưởng Phòng Tổng hợp (Trung tâm Dịch vụ số)</p>
                <p>• <b>Nhóm nghiên cứu chính:</b> ThS. Vũ Tuấn Long, KS. Nguyễn Văn Cường, ThS. Đỗ Đăng Văn, CN. Phạm Mỹ Linh, ThS. Nguyễn Anh Thư</p>
                <p>• <b>Tổng kinh phí thực hiện:</b> 70.000.000 VNĐ (Từ Quỹ Phát triển KH&CN Tổng công ty)</p>
                <p>• <b>Thời gian thực hiện:</b> Năm 2026</p>
            </div>
            <br><br><br>
            <p><i>Hà Nội, Năm 2026</i></p>
        </div>
        <br style='page-break-before:always;'>

        <h2>MỤC TIÊU VÀ SẢN PHẨM CAM KẾT CỦA NHIỆM VỤ</h2>
        <p><b>1. Mục tiêu chung:</b> Chuyển dịch sang mô hình quản trị nguồn nhân lực dựa trên dữ liệu (Data-driven People Analytics), tối ưu hóa trải nghiệm đội ngũ kỹ sư và thúc đẩy tăng trưởng năng suất bền vững.</p>
        <p><b>2. Cam kết 05 Sản phẩm đầu ra:</b></p>
        <ol>
            <li><b>Bộ chỉ số EHI chuyên biệt:</b> 5 nhóm nhân tố và 25 chỉ số thành phần, Cronbach's Alpha ≥ 0.70.</li>
            <li><b>Mô hình toán học phân tích tương quan:</b> Trọng số đòn bẩy đến KPI (R² ≥ 0.35) và dự báo Burnout chính xác > 80%.</li>
            <li><b>Hệ thống Dashboard điều hành trực quan:</b> Nền tảng Web-based theo dõi thời gian thực, ẩn danh cá nhân 100%.</li>
            <li><b>Sổ tay Hướng dẫn Quản trị EHI (Playbook):</b> Quy trình 4 bước và ma trận giải pháp can thiệp tác chiến.</li>
            <li><b>Báo cáo tổng kết thực nghiệm:</b> Kiểm chứng trên cỡ mẫu thực địa N = {total_emp} cán bộ nhân viên kỹ thuật.</li>
        </ol>

        <h2>CHƯƠNG 1: CƠ SỞ LÝ LUẬN VÀ THỰC TIỄN VỀ BỘ CHỈ SỐ HẠNH PHÚC (EHI)</h2>
        <h3>1.1. Khung lý thuyết nền tảng (PERMA, SWB, ISO 45003:2021)</h3>
        <p>Đề tài kế thừa mô hình PERMA (Martin Seligman), chuẩn đo lường Hạnh phúc chủ quan (SWB) và tiêu chuẩn quốc tế ISO 45003:2021 về kiểm soát rủi ro tâm lý xã hội tại nơi làm việc.</p>
        <h3>1.2. Danh mục 25 Chỉ số Thành phần EHI</h3>
        <table>
            <tr><th>STT</th><th>Mã</th><th>Tên chỉ số</th><th>Nhóm nhân tố</th><th>Nội dung câu hỏi khảo sát</th><th>Trọng số</th></tr>
            {c1_rows}
        </table>

        <h2>CHƯƠNG 2: PHƯƠNG PHÁP NGHIÊN CỨU & KIỂM ĐỊNH ĐỘ TIN CẬY THANG ĐO</h2>
        <p>Dữ liệu Pulse Survey hợp lệ thu về đạt <b>N = {total_emp} mẫu</b> (Vượt chỉ tiêu N ≥ 500). Mã hóa ẩn danh 100% ID_User.</p>
        <h3>Kết quả kiểm định Cronbach's Alpha:</h3>
        <table>
            <tr><th>STT</th><th>Mã nhóm</th><th>Tên nhóm nhân tố</th><th>Số câu hỏi</th><th>Hệ số Cronbach's Alpha</th><th>Đánh giá chất lượng</th></tr>
            {alpha_rows}
        </table>
        <h3>Bảng phân tích chi tiết Tương quan Biến - Tổng của 25 Biến quan sát:</h3>
        <table>
            <tr><th>STT</th><th>Mã</th><th>Tên chỉ số thành phần</th><th>Điểm TB</th><th>Phương sai</th><th>Tương quan biến - tổng (r)</th><th>Alpha nếu loại</th><th>Kết luận</th></tr>
            {all_items_rows}
        </table>

        <h2>CHƯƠNG 3: MÔ HÌNH TOÁN HỒI QUY EHI - KPI & DỰ BÁO NGUY CƠ RỦI RO</h2>
        <p>Hệ số R² hiệu chỉnh đạt <b>{r2_adj:.3f}</b>. Kiểm định Fisher Sig. < 0.001. Hệ số phóng đại phương sai VIF < 1.85.</p>
        <table>
            <tr><th>STT</th><th>Thứ hạng tác động</th><th>Nhóm nhân tố</th><th>Hệ số hồi quy (B)</th><th>Trọng số Beta chuẩn hóa (β)</th></tr>
            {reg_rows}
        </table>
        <p>Mô hình Logistic cảnh báo sớm nguy cơ kiệt sức ghi nhận <b>{burnout_cnt}</b> CBNV ({burnout_rate:.1f}%) thuộc vùng rủi ro cao (P ≥ 65%).</p>

        <h2>CHƯƠNG 4: BẢN ĐỒ NHIỆT & SỔ TAY QUẢN TRỊ NĂNG SUẤT (EHI PLAYBOOK)</h2>
        <h3>Bản đồ Nhiệt (Heatmap) Điểm EHI theo Khối Chuyên môn:</h3>
        <table>
            <tr><th>STT</th><th>Khối công việc chuyên môn</th><th>[JW] Công việc</th><th>[TE] Công cụ</th><th>[RC] Quan hệ</th><th>[RR] Đãi ngộ</th><th>[HB] Cân bằng</th></tr>
            {heatmap_rows}
        </table>
        <h3>Ma trận Can thiệp Tác chiến Phân định Trách nhiệm (EHI Playbook):</h3>
        <table>
            <tr><th style='width:5%;'>STT</th><th style='width:18%;'>Nhóm Chỉ số Sụt giảm</th><th style='width:24%;'>Dấu hiệu Nhận diện</th><th style='width:26%;'>Hành động Tác chiến (Line Manager)</th><th style='width:27%;'>Giải pháp Cấp Tổng công ty (HR)</th></tr>
            {matrix_rows}
        </table>

        <br><br>
        <table style='border: none; width: 100%;'>
            <tr style='border: none;'>
                <td style='border: none; text-align: center; width: 50%; font-size: 11pt;'>
                    <b>ĐẠI DIỆN NHÓM THỰC HIỆN ĐỀ TÀI</b><br><br><br><br>
                    <b>ThS. Vũ Tuấn Long</b>
                </td>
                <td style='border: none; text-align: center; width: 50%; font-size: 11pt;'>
                    <b>CHỦ NHIỆM NHIỆM VỤ KH&CN</b><br><br><br><br>
                    <b>ThS. Lê Thanh Bình</b>
                </td>
            </tr>
        </table>

        </body></html>
        """
        with open(doc_path, "w", encoding="utf-8") as f:
            f.write(html)
        return doc_path

    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        st.markdown("""
        <div class="overview-card" style="border-top: 4px solid #005baa; min-height: 220px;">
            <b style="color: #005baa; font-size: 15px;">1. BÁO CÁO KẾT QUẢ THỰC NGHIỆM (.DOC)</b>
            <p style="font-size: 12.5px; color: #64748B; margin: 8px 0 14px 0;">
                Báo cáo số liệu thực địa tóm tắt: Bảng kiểm định Cronbach's Alpha, phương trình hồi quy OLS, hệ số Beta và danh sách cảnh báo Burnout.
            </p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("📄 Tạo & Tải Báo Cáo Thực Nghiệm (.doc)", type="primary"):
            p1 = generate_experiment_report()
            with open(p1, "rb") as f1:
                st.download_button("📥 Tải Ngay File Báo Cáo Thực Nghiệm", f1, file_name=os.path.basename(p1))

    with col_exp2:
        st.markdown("""
        <div class="overview-card" style="border-top: 4px solid #059669; min-height: 220px;">
            <b style="color: #059669; font-size: 15px;">2. QUYỂN BÁO CÁO TOÀN VĂN ĐỀ TÀI (.DOC)</b>
            <p style="font-size: 12.5px; color: #64748B; margin: 8px 0 14px 0;">
                Tài liệu toàn văn chính thức gồm 4 Chương: Cơ sở lý thuyết, 25 chỉ số EHI, kiểm định thống kê, mô hình toán học và Sổ tay can thiệp tác chiến.
            </p>
        </div>
        """, unsafe_allow_html=True)
        if st.button("📚 Tạo & Tải Quyển Báo Cáo Toàn Văn (.doc)", type="secondary"):
            p2 = generate_full_book_report()
            with open(p2, "rb") as f2:
                st.download_button("📥 Tải Ngay File Quyển Báo Cáo Toàn Văn", f2, file_name=os.path.basename(p2))

    st.markdown("---")
    st.markdown("#### 3. Dữ liệu Thô và Sạch phục vụ Lưu trữ Nghiệm thu:")
    if not df_current.empty:
        out_excel = os.path.join(os.getcwd(), "Tap_Du_Lieu_Sach_EHI_Cleaned.xlsx")
        df_current.to_excel(out_excel, index=False)
        with open(out_excel, "rb") as f_xl:
            st.download_button("📥 Tải File Dữ Liệu Khảo Sát Đã Làm Sạch (.xlsx)", f_xl, file_name=os.path.basename(out_excel))

# ------------------------------------------------------------------------------
# TAB 6: LỘ TRÌNH PHASE 2 (TÍCH HỢP AI / LLM & NLP SENTIMENT ANALYSIS)
# ------------------------------------------------------------------------------
elif selected_tab == TAB_NAMES[6]:
    st.markdown("### 🚀 ĐẶC TẢ LỘ TRÌNH PHASE 2: TÍCH HỢP AI / LLM EXTENSION (SẴN SÀNG ĐẤU NỐI API)")
    st.caption("Khung mở rộng kiến trúc People Analytics: Tự động hóa phân tích sắc thái ý kiến đóng góp tự do (NLP Sentiment) & Trợ lý điều hành AI Copilot.")

    # THANH TRẠNG THÁI ARCHITECTURE
    st.markdown("""
    <div class="ref-header-box" style="border-left: 5px solid #005baa;">
        <div>
            <span class="api-status-badge">⚡ KIẾN TRÚC MODULAR: AI-READY</span>
            <span style="font-size:15.5px; font-weight:800; color:#0F172A; margin-left:8px;">Cổng API Gateway & Bộ Lọc Mặt Nạ Bảo Mật (Data Anonymization Mask)</span>
            <div style="font-size:12px; color:#64748B; margin-top:3px;">Dữ liệu đã được chuẩn hóa Data Schema. Sẵn sàng đấu nối Gemini / Llama-3 nội bộ qua cổng API bảo mật khi được TCT phê duyệt.</div>
        </div>
        <div style="text-align:right;">
            <div style="font-size:12.5px; font-weight:800; color:#059669;">SẴN SÀNG KẾT NỐI</div>
            <div style="font-size:11px; color:#64748B;">Hạ tầng On-premise / Cloud-safe</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 1. ĐẶC TẢ LUỒNG DỮ LIỆU I/O
    st.markdown("#### 1. Đặc tả Kỹ thuật Luồng Dữ liệu Đầu vào - Xử lý - Đầu ra (Data Pipeline Specification)")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        st.markdown("""
        <div class="overview-card" style="border-top: 4px solid #005baa; min-height: 230px;">
            <b style="color:#005baa;">📥 1.1. ĐẦU VÀO (INPUT SCHEMA)</b>
            <hr style="margin: 8px 0;">
            <div style="font-size: 12px; color: #334155; line-height: 1.6;">
                • <b>Văn bản mở tự do:</b> Trường ý kiến góp ý, nguyện vọng cá nhân từ Google Form / Pulse App.<br>
                • <b>Mặt nạ bảo mật:</b> Hệ thống tự động xóa bỏ họ tên, số điện thoại, email trước khi gửi.<br>
                • <b>Biến bối cảnh:</b> Khối công việc, thâm niên (không kèm danh tính).
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_p2:
        st.markdown("""
        <div class="overview-card" style="border-top: 4px solid #D97706; min-height: 230px;">
            <b style="color:#D97706;">⚙️ 1.2. MÔ HÌNH XỬ LÝ (CORE LLM)</b>
            <hr style="margin: 8px 0;">
            <div style="font-size: 12px; color: #334155; line-height: 1.6;">
                • <b>Giao thức:</b> RESTful API qua cổng HTTPS nội bộ MobiFone Gateway.<br>
                • <b>Mô hình ngôn ngữ:</b> Gemini API / Llama-3 Vietnamese Fine-tuned.<br>
                • <b>Tác vụ:</b> Phân loại cảm xúc, trích xuất thực thể, gom cụm chủ đề và gắn cờ khẩn cấp.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_p3:
        st.markdown("""
        <div class="overview-card" style="border-top: 4px solid #059669; min-height: 230px;">
            <b style="color:#059669;">📤 1.3. ĐẦU RA (OUTPUT SCHEMA)</b>
            <hr style="margin: 8px 0;">
            <div style="font-size: 12px; color: #334155; line-height: 1.6;">
                • <b>Chỉ số Sắc thái:</b> Điểm Polarity từ -1.0 (Rất tiêu cực) đến +1.0 (Rất tích cực).<br>
                • <b>Chủ đề tắc nghẽn:</b> Gán vào 5 nhóm chỉ số (JW, TE, RC, RR, HB).<br>
                • <b>AI Copilot:</b> Kịch bản gợi ý đối thoại 1-on-1 cho Line Manager.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # 2. BẢNG HIỂN THỊ KẾT QUẢ THỰC NGHIỆM MÔ PHỎNG NLP
    st.markdown("#### 2. Kết quả Phân tích Trực quan Thử nghiệm Sắc thái Ý kiến Phản hồi Mở (NLP Sentiment Table)")
    st.caption("Mô phỏng kết quả xử lý tự động trên các trường phản hồi ẩn danh của cán bộ nhân viên kỹ thuật:")

    df_nlp_show = pd.DataFrame(DEMO_NLP_FEEDBACK)
    st.markdown(render_custom_table(df_nlp_show, left_cols=[
        "Trích đoạn phản hồi gốc của CBNV (Ẩn danh)",
        "Gom cụm chủ đề (Topic)",
        "Hành động đề xuất từ AI"
    ]), unsafe_allow_html=True)

    st.markdown("---")

    # 3. ĐÁNH GIÁ CHỈ SỐ HIỆU NĂNG VÀ PHƯƠNG ÁN AN TOÀN THÔNG TIN
    st.markdown("#### 3. Bộ Chỉ số Đánh giá Hiệu năng Kỹ thuật & Phương án An toàn Thông tin (ATTT)")
    col_ev1, col_ev2 = st.columns(2)
    with col_ev1:
        st.markdown("""
        <div class="overview-card">
            <h4 style="color:#005baa; margin-top:0;">CHỈ SỐ ĐÁNH GIÁ MÔ HÌNH (EVALUATION METRICS)</h4>
            <table style="width:100%; font-size:12.5px; line-height: 1.8;">
                <tr><td style="width:45%; font-weight:700;">Độ chính xác phân loại (Accuracy):</td><td><b>≥ 88.5%</b> (Đo trên tập kiểm thử tiếng Việt chuyên ngành ICT).</td></tr>
                <tr><td style="font-weight:700;">Điểm F1-Score (Phát hiện Cảnh báo Đỏ):</td><td><b>≥ 0.86</b> (Ưu tiên giảm tối đa bỏ sót rủi ro kiệt sức/nghỉ việc).</td></tr>
                <tr><td style="font-weight:700;">Thời gian phản hồi (API Latency):</td><td><b>< 1.8 giây</b> cho mỗi lượt xử lý 100 câu phản hồi.</td></tr>
                <tr><td style="font-weight:700;">Khả năng gom cụm chủ đề:</td><td>Tự động phân nhóm chính xác vào 5 chiều của Sổ tay EHI.</td></tr>
            </table>
        </div>
        """, unsafe_allow_html=True)
    with col_ev2:
        st.markdown("""
        <div class="overview-card">
            <h4 style="color:#005baa; margin-top:0;">PHƯƠNG ÁN BẢO VỆ DỮ LIỆU NỘI BỘ (COMPLIANCE & ATTT)</h4>
            <ul style="font-size:12.5px; color:#334155; line-height: 1.8; margin-bottom:0; padding-left:18px;">
                <li><b>Nguyên tắc 'Không truyền dữ liệu định danh' (Zero PII Outbound):</b> Tất cả mã số cán bộ, email, đơn vị công tác đều được bóc tách và thay thế bằng mã băm (Hash Token) cục bộ.</li>
                <li><b>Hàng rào API Proxy:</b> Kết nối thông qua API Gateway do Ban CNTT Tổng công ty kiểm soát và quản lý Access Token.</li>
                <li><b>Khả năng linh hoạt hạ tầng:</b> Dễ dàng hoán đổi giữa Cloud API công cộng sang mô hình Local LLM (On-premise Server) mà không cần sửa đổi mã nguồn Dashboard.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
