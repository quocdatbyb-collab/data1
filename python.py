#import streamlit as st
import pandas as pd
from google import genai
from google.genai.errors import APIError

# --- Cấu hình Trang Streamlit ---
st.set_page_config(
    page_title="App Phân Tích Báo Cáo Tài Chính",
    layout="wide"
)

st.title("Ứng dụng Phân Tích Báo Cáo Tài chính 📊")

# --- Hàm tính toán chính (Sử dụng Caching để Tối ưu hiệu suất) ---
@st.cache_data
def process_financial_data(df):
    """Thực hiện các phép tính Tăng trưởng và Tỷ trọng."""
    
    # Đảm bảo các giá trị là số để tính toán
    numeric_cols = ['Năm trước', 'Năm sau']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # 1. Tính Tốc độ Tăng trưởng
    # Dùng .replace(0, 1e-9) cho Series Pandas để tránh lỗi chia cho 0
    df['Tốc độ tăng trưởng (%)'] = (
        (df['Năm sau'] - df['Năm trước']) / df['Năm trước'].replace(0, 1e-9)
    ) * 100

    # 2. Tính Tỷ trọng theo Tổng Tài sản
    # Lọc chỉ tiêu "TỔNG CỘNG TÀI SẢN"
    tong_tai_san_row = df[df['Chỉ tiêu'].str.contains('TỔNG CỘNG TÀI SẢN', case=False, na=False)]
    
    if tong_tai_san_row.empty:
        raise ValueError("Không tìm thấy chỉ tiêu 'TỔNG CỘNG TÀI SẢN'.")

    tong_tai_san_N_1 = tong_tai_san_row['Năm trước'].iloc[0]
    tong_tai_san_N = tong_tai_san_row['Năm sau'].iloc[0]

    # ******************************* PHẦN SỬA LỖI BẮT ĐẦU *******************************
    # Lỗi xảy ra khi dùng .replace() trên giá trị đơn lẻ (numpy.int64).
    # Sử dụng điều kiện ternary để xử lý giá trị 0 thủ công cho mẫu số.
    
    divisor_N_1 = tong_tai_san_N_1 if tong_tai_san_N_1 != 0 else 1e-9
    divisor_N = tong_tai_san_N if tong_tai_san_N != 0 else 1e-9

    # Tính tỷ trọng với mẫu số đã được xử lý
    df['Tỷ trọng Năm trước (%)'] = (df['Năm trước'] / divisor_N_1) * 100
    df['Tỷ trọng Năm sau (%)'] = (df['Năm sau'] / divisor_N) * 100
    # ******************************* PHẦN SỬA LỖI KẾT THÚC *******************************
    
    return df

# --- Hàm gọi API Gemini cho Nhận xét Tự động (Chức năng 5) ---
# Hàm này được giữ nguyên, chỉ dùng cho nút "Yêu cầu AI Phân tích"
def get_ai_analysis(data_for_ai, api_key):
    """Gửi dữ liệu phân tích đến Gemini API và nhận nhận xét."""
    try:
        client = genai.Client(api_key=api_key)
        model_name = 'gemini-2.5-flash' 

        prompt = f"""
        Bạn là một chuyên gia phân tích tài chính chuyên nghiệp. Dựa trên các chỉ số tài chính sau, hãy đưa ra một nhận xét khách quan, ngắn gọn (khoảng 3-4 đoạn) về tình hình tài chính của doanh nghiệp. Đánh giá tập trung vào tốc độ tăng trưởng, thay đổi cơ cấu tài sản và khả năng thanh toán hiện hành.
        
        Dữ liệu thô và chỉ số:
        {data_for_ai}
        """

        response = client.models.generate_content(
            model=model_name,
            contents=prompt
        )
        return response.text

    except APIError as e:
        return f"Lỗi gọi Gemini API: Vui lòng kiểm tra Khóa API hoặc giới hạn sử dụng. Chi tiết lỗi: {e}"
    except KeyError:
        return "Lỗi: Không tìm thấy Khóa API 'GEMINI_API_KEY'. Vui lòng kiểm tra cấu hình Secrets trên Streamlit Cloud."
    except Exception as e:
        return f"Đã xảy ra lỗi không xác định: {e}"


# --- Chức năng 1: Tải File ---
uploaded_file = st.file_uploader(
    "1. Tải file Excel Báo cáo Tài chính (Chỉ tiêu | Năm trước | Năm sau)",
    type=['xlsx', 'xls']
)

# Khởi tạo các biến để đảm bảo luôn tồn tại (dùng cho chat context)
thanh_toan_hien_hanh_N = "N/A"
thanh_toan_hien_hanh_N_1 = "N/A"
data_for_ai = ""

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        # Tiền xử lý: Đảm bảo chỉ có 3 cột quan trọng
        df_raw.columns = ['Chỉ tiêu', 'Năm trước', 'Năm sau']
        
        # Xử lý dữ liệu
        df_processed = process_financial_data(df_raw.copy())

        if df_processed is not None:
            
            # --- Chức năng 2 & 3: Hiển thị Kết quả ---
            st.subheader("2. Tốc độ Tăng trưởng & 3. Tỷ trọng Cơ cấu Tài sản")
            st.dataframe(df_processed.style.format({
                'Năm trước': '{:,.0f}',
                'Năm sau': '{:,.0f}',
                'Tốc độ tăng trưởng (%)': '{:.2f}%',
                'Tỷ trọng Năm trước (%)': '{:.2f}%',
                'Tỷ trọng Năm sau (%)': '{:.2f}%'
            }), use_container_width=True)
            
            # --- Chức năng 4: Tính Chỉ số Tài chính ---
            st.subheader("4. Các Chỉ số Tài chính Cơ bản")
            
            try:
                # Lọc giá trị cho Chỉ số Thanh toán Hiện hành (Ví dụ)
                
                # Lấy Tài sản ngắn hạn
                tsnh_n = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Năm sau'].iloc[0]
                tsnh_n_1 = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Năm trước'].iloc[0]

                # Lấy Nợ ngắn hạn (Dùng giá trị giả định hoặc lọc từ file nếu có)
                # **LƯU Ý: Thay thế logic sau nếu bạn có Nợ Ngắn Hạn trong file**
                no_ngan_han_N = df_processed[df_processed['Chỉ tiêu'].str.contains('NỢ NGẮN HẠN', case=False, na=False)]['Năm sau'].iloc[0]  
                no_ngan_han_N_1 = df_processed[df_processed['Chỉ tiêu'].str.contains('NỢ NGẮN HẠN', case=False, na=False)]['Năm trước'].iloc[0]

                # Tính toán
                thanh_toan_hien_hanh_N = tsnh_n / no_ngan_han_N
                thanh_toan_hien_hanh_N_1 = tsnh_n_1 / no_ngan_han_N_1
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric(
                        label="Chỉ số Thanh toán Hiện hành (Năm trước)",
                        value=f"{thanh_toan_hien_hanh_N_1:.2f} lần"
                    )
                with col2:
                    st.metric(
                        label="Chỉ số Thanh toán Hiện hành (Năm sau)",
                        value=f"{thanh_toan_hien_hanh_N:.2f} lần",
                        delta=f"{thanh_toan_hien_hanh_N - thanh_toan_hien_hanh_N_1:.2f}"
                    )
                    
            except IndexError:
                 st.warning("Thiếu chỉ tiêu 'TÀI SẢN NGẮN HẠN' hoặc 'NỢ NGẮN HẠN' để tính chỉ số.")
                 thanh_toan_hien_hanh_N = "N/A" # Dùng để tránh lỗi ở Chức năng 5
                 thanh_toan_hien_hanh_N_1 = "N/A"
            
            # --- Chức năng 5: Nhận xét AI ---
            st.subheader("5. Nhận xét Tình hình Tài chính (AI)")
            
            # Chuẩn bị dữ liệu để gửi cho AI (được dùng chung cho Chat Context)
            data_for_ai = pd.DataFrame({
                'Chỉ tiêu': [
                    'Toàn bộ Bảng phân tích (dữ liệu thô)', 
                    'Tăng trưởng Tài sản ngắn hạn (%)', 
                    'Thanh toán hiện hành (N-1)', 
                    'Thanh toán hiện hành (N)'
                ],
                'Giá trị': [
                    df_processed.to_markdown(index=False),
                    f"{df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Tốc độ tăng trưởng (%)'].iloc[0]:.2f}%" if thanh_toan_hien_hanh_N != "N/A" else "N/A", 
                    f"{thanh_toan_hien_hanh_N_1}", 
                    f"{thanh_toan_hien_hanh_N}"
                ]
            }).to_markdown(index=False) 

            if st.button("Yêu cầu AI Phân tích"):
                api_key = st.secrets.get("GEMINI_API_KEY") 
                
                if api_key:
                    with st.spinner('Đang gửi dữ liệu và chờ Gemini phân tích...'):
                        ai_result = get_ai_analysis(data_for_ai, api_key)
                        st.markdown("**Kết quả Phân tích từ Gemini AI:**")
                        st.info(ai_result)
                else:
                     st.error("Lỗi: Không tìm thấy Khóa API. Vui lòng cấu hình Khóa 'GEMINI_API_KEY' trong Streamlit Secrets.")

            
            # *************************************************************************
            # ************************ BẮT ĐẦU CHỨC NĂNG CHAT *************************
            # *************************************************************************

            st.markdown("---")
            st.subheader("6. Chat Hỏi đáp Chuyên sâu với Gemini")
            
            # --- 1. Khởi tạo State Chat ---
            # gemini_chat_history sẽ lưu trữ toàn bộ cuộc trò chuyện (bao gồm system prompt ẩn)
            if "gemini_chat_history" not in st.session_state:
                st.session_state.gemini_chat_history = []
            
            # Chỉ khởi tạo context khi file được tải lên mới (hoặc lần đầu)
            if "last_uploaded_filename" not in st.session_state or st.session_state.last_uploaded_filename != uploaded_file.name:
                st.session_state.gemini_chat_history = [] # Reset lịch sử
                st.session_state.last_uploaded_filename = uploaded_file.name
                
                # System Instruction để grounding model với dữ liệu
                system_prompt = f"""
                Bạn là một chuyên gia phân tích tài chính am hiểu và hỗ trợ chat. Nhiệm vụ của bạn là trả lời các câu hỏi của người dùng dựa trên Bảng Dữ liệu Tài chính đã được xử lý sau đây.
                Hãy sử dụng **dữ liệu này** làm nền tảng cho mọi phân tích và giải thích của bạn.

                **BẢNG DỮ LIỆU TÀI CHÍNH ĐÃ XỬ LÝ VÀ CHỈ SỐ QUAN TRỌNG:**
                {data_for_ai}
                
                Không cần nhắc lại dữ liệu này trong mỗi câu trả lời, hãy tập trung vào câu hỏi của người dùng.
                """
                
                # Thêm System Instruction vào lịch sử chat (không hiển thị)
                st.session_state.gemini_chat_history.append({"role": "system", "content": system_prompt})
                # Tin nhắn chào mừng hiển thị cho người dùng
                st.session_state.gemini_chat_history.append({"role": "assistant", "content": "Dữ liệu đã sẵn sàng. Tôi có thể giúp bạn giải đáp các thắc mắc chuyên sâu nào về tăng trưởng, cơ cấu hay các chỉ số tài chính của doanh nghiệp?"})

            # --- 2. Hiển thị lịch sử chat ---
            api_key = st.secrets.get("GEMINI_API_KEY")
            client = None
            
            # Khởi tạo client một lần
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                except Exception:
                    # Bỏ qua lỗi khởi tạo, sẽ xử lý khi người dùng gửi tin nhắn
                    pass

            # Hiển thị tất cả tin nhắn trừ "system"
            for message in st.session_state.gemini_chat_history:
                if message["role"] != "system":
                    with st.chat_message(message["role"]):
                        st.markdown(message["content"])

            # --- 3. Xử lý input từ người dùng ---
            if prompt := st.chat_input("Hỏi Gemini về Báo cáo Tài chính của bạn..."):
                
                # 3.1. Hiển thị prompt của người dùng
                with st.chat_message("user"):
                    st.markdown(prompt)
                
                # 3.2. Cập nhật state
                st.session_state.gemini_chat_history.append({"role": "user", "content": prompt})

                if not api_key:
                    st.error("Lỗi: Không tìm thấy Khóa API 'GEMINI_API_KEY'. Vui lòng cấu hình Streamlit Secrets.")
                    # Xóa tin nhắn người dùng vừa nhập nếu lỗi API
                    st.session_state.gemini_chat_history.pop() 
                    st.rerun()

                if client is None:
                    st.error("Lỗi khởi tạo Client API. Vui lòng kiểm tra Khóa API.")
                    # Xóa tin nhắn người dùng vừa nhập nếu lỗi Client
                    st.session_state.gemini_chat_history.pop() 
                    st.rerun()
                    
                with st.chat_message("assistant"):
                    with st.spinner("Gemini đang phân tích và trả lời..."):
                        try:
                            # Lấy toàn bộ lịch sử chat (bao gồm cả system prompt)
                            # Chuyển đổi định dạng lịch sử Streamlit sang định dạng API
                            contents = [{"role": m["role"], "parts": [{"text": m["content"]}]} 
                                        for m in st.session_state.gemini_chat_history]
                            
                            # Lấy response từ Gemini
                            response = client.models.generate_content(
                                model='gemini-2.5-flash',
                                contents=contents
                            )

                            ai_response = response.text
                            st.markdown(ai_response)
                            
                            # Cập nhật state với câu trả lời của AI
                            st.session_state.gemini_chat_history.append({"role": "assistant", "content": ai_response})

                        except APIError as e:
                            error_message = f"Lỗi gọi Gemini API: Vui lòng kiểm tra Khóa API hoặc giới hạn sử dụng. Chi tiết lỗi: {e}"
                            st.error(error_message)
                            st.session_state.gemini_chat_history.pop() # Remove user prompt on error
                        except Exception as e:
                            error_message = f"Đã xảy ra lỗi không xác định: {e}"
                            st.error(error_message)
                            st.session_state.gemini_chat_history.pop() # Remove user prompt on error
            
            # ************************* KẾT THÚC CHỨC NĂNG CHAT ***********************

    except ValueError as ve:
        st.error(f"Lỗi cấu trúc dữ liệu: {ve}")
    except Exception as e:
        st.error(f"Có lỗi xảy ra khi đọc hoặc xử lý file: {e}. Vui lòng kiểm tra định dạng file.")

else:
    st.info("Vui lòng tải lên file Excel để bắt đầu phân tích.")
