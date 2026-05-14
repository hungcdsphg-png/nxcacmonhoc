
import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";
import { StudentRecord } from "../types";

const SYSTEM_INSTRUCTION = `
Bạn là trợ lý viết nhận xét học bạ tiểu học (Thông tư 27). 
Nhiệm vụ: Viết nhận xét ngắn gọn (khoảng 150 ký tự), dễ hiểu, mộc mạc cho học sinh vùng dân tộc thiểu số.

QUY TẮC NGÔN NGỮ (BẮT BUỘC):
- TUYỆT ĐỐI KHÔNG dùng từ: "con", "em", "bé", "thầy", "cô", "thầy giáo", "cô giáo".
- TUYỆT ĐỐI KHÔNG dùng từ: "bản", "làng", "bản làng".
- TUYỆT ĐỐI KHÔNG dùng tên riêng của học sinh.
- Sử dụng tiếng Việt phổ thông đơn giản, không dùng từ địa phương, không dùng thuật ngữ sư phạm hàn lâm.

QUY TẮC NỘI DUNG DỰA TRÊN PPCT (BẮT BUỘC):
- Đối với các môn học (ngoại trừ Tiếng Việt): Phải sử dụng các tên bài học, chủ đề, bài đọc, nội dung viết cụ thể từ danh sách PPCT đã cung cấp.
- ĐẶC BIỆT CHO MÔN TIẾNG VIỆT: TUYỆT ĐỐI KHÔNG ghi tên bài học, tên bài đọc cụ thể vào nhận xét. Hãy tập trung nhận xét về các kỹ năng (Đọc, Viết, Nói và Nghe) dựa trên mức đạt được và điểm số.
- Nhận xét phải phản ánh đúng kiến thức của học kỳ đang chọn (Giữa kì 1, Cuối kì 1, Giữa kì 2, Cuối kì 2).
- Ví dụ (Môn khác): Nếu PPCT có bài "Bầu trời", nhận xét điểm 10 có thể là: "Đọc to, rõ ràng bài Bầu trời. Hiểu nội dung bài và trả lời đúng các câu hỏi. Viết chữ đẹp, đúng độ cao."
- Ví dụ (Tiếng Việt): "Đọc to, rõ ràng, trôi chảy, hiểu nội dung đoạn văn. Viết chữ đúng mẫu, đều nét. Biết cách dùng từ đặt câu linh hoạt."

QUY TẮC PHÂN LOẠI THEO ĐIỂM (PHẢI TUÂN THỦ NGHIÊM NGẶT):
- Điểm 10: Mức T. Nhận xét: Hoàn thành xuất sắc, nắm vững các bài đọc và nội dung viết trong học kỳ. Trình bày khoa học, sáng tạo.
- Điểm 9: Mức T. Nhận xét: Hoàn thành rất tốt các chủ đề học tập, tự giác cao, bài làm cẩn thận, đúng yêu cầu.
- Điểm 8: Mức T. Nhận xét: Hoàn thành tốt nội dung môn học, tích cực phát biểu, nắm chắc kiến thức các bài đã học.
- Điểm 7: Mức H. Nhận xét: Hoàn thành khá tốt các yêu cầu, nắm được kiến thức trọng tâm nhưng đôi khi còn thiếu cẩn thận khi viết.
- Điểm 6: Mức H. Nhận xét: Hoàn thành nội dung cơ bản, nắm được nội dung các bài đọc nhưng còn lúng túng ở phần viết/luyện tập.
- Điểm 5: Mức H. Nhận xét: Hoàn thành mức độ vừa đủ, kiến thức cơ bản về các chủ đề còn chưa chắc chắn, cần nỗ lực luyện tập thêm.
- Điểm 4: Mức C. Nhận xét: Chưa hoàn thành một số bài học, kiến thức còn hổng nhiều, cần được kèm cặp sát sao các phần Đọc/Viết.
- Điểm 3: Mức C. Nhận xét: Chưa nắm được kiến thức cơ bản của học kỳ, kết quả học tập còn hạn chế, cần tập trung và đi học đều hơn.

YÊU CẦU VỀ SỰ KHÁC BIỆT:
- Nhận xét cho điểm 10 PHẢI khác biệt và cao cấp hơn điểm 9.
- Nhận xét cho điểm 9 PHẢI tốt hơn điểm 8.
- Tương tự cho các mức điểm khác. Không được dùng chung một mẫu nhận xét cho các mức điểm khác nhau.

VĂN PHONG MẪU: "Đọc to, rõ ràng. Làm toán đúng và nhanh. Chăm chỉ học tập, tích cực phát biểu xây dựng bài. Cần giữ vững tinh thần học tập này."
`;

export interface BankComment {
  id: string;
  mucDo: 'T' | 'H' | 'C';
  diem: number;
  noiDung: string;
}

// Helper function for exponential backoff retry
const callWithRetry = async (fn: () => Promise<any>, maxRetries = 3, initialDelay = 2000) => {
  let retries = 0;
  while (retries < maxRetries) {
    try {
      return await fn();
    } catch (error: any) {
      // Extract error details from Gemini API response
      const errorData = error?.error || error;
      const isRateLimit = errorData?.code === 429 || errorData?.status === "RESOURCE_EXHAUSTED" || error?.message?.includes("429");
      
      if (isRateLimit) {
        if (retries < maxRetries - 1) {
          const delay = initialDelay * Math.pow(2, retries);
          console.warn(`Rate limit hit (429). Retrying in ${delay}ms... (Attempt ${retries + 1}/${maxRetries})`);
          await new Promise(resolve => setTimeout(resolve, delay));
          retries++;
          continue;
        }
        throw new Error("QUOTA_EXCEEDED");
      }
      throw error;
    }
  }
};

export const extractLessonsFromPpct = async (
  rawText: string,
  subject: string,
  gradeLevel: string,
  apiKey: string
): Promise<string> => {
  const ai = new GoogleGenAI({ apiKey });
  
  // Tiền xử lý văn bản thô để giảm nhiễu và tăng tốc độ
  const preProcessedText = rawText
    .split('\n')
    .filter(line => line.trim().length > 5) // Loại bỏ các dòng quá ngắn
    .slice(0, 800) // Tăng giới hạn lên 800 dòng để lấy đủ dữ liệu nhưng vẫn nhanh
    .join('\n');

  const prompt = `Dưới đây là nội dung trích xuất từ file Phân phối chương trình (PPCT) môn ${subject}, ${gradeLevel}.
  Nhiệm vụ: Trích xuất danh sách các bài học và nội dung chi tiết một cách CHÍNH XÁC và NHANH CHÓNG.
  
  QUY TẮC TRÍCH XUẤT:
  1. BỎ QUA: Số tuần, số tiết, ngày tháng, ghi chú, tên giáo viên, tên trường.
  2. TẬP TRUNG: Tên bài học (Lesson title) và Yêu cầu cần đạt/Nội dung chính/Hoạt động.
  3. TRẢ VỀ JSON: Danh sách các đối tượng có "title" (Tên bài) và "details" (mảng các hoạt động/nội dung chi tiết).
  
  NỘI DUNG CẦN PHÂN TÍCH:
  "${preProcessedText}"`;

  try {
    const response = await callWithRetry(() => ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
      config: {
        systemInstruction: "Bạn là chuyên gia trích xuất dữ liệu giáo dục. Hãy làm việc cực nhanh và chính xác. Chỉ trả về JSON, không giải thích.",
        thinkingConfig: { thinkingLevel: ThinkingLevel.MINIMAL },
        temperature: 0.1,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              title: { type: Type.STRING, description: "Tên bài học hoặc chủ đề" },
              details: { 
                type: Type.ARRAY, 
                items: { type: Type.STRING },
                description: "Các hoạt động chi tiết như Đọc, Viết, Luyện tập..."
              }
            },
            required: ["title", "details"]
          }
        }
      },
    }));

    const lessons = JSON.parse(response.text || "[]");
    
    // Format lại thành chuỗi mà UI mong đợi
    if (Array.isArray(lessons) && lessons.length > 0) {
      return lessons.map((lesson: any, index: number) => {
        let text = `BÀI ${index + 1}: ${lesson.title.toUpperCase()}\n`;
        if (lesson.details && Array.isArray(lesson.details)) {
          lesson.details.forEach((detail: string) => {
            // Đảm bảo có dấu hai chấm cho UI nhận diện label
            if (!detail.includes(':')) {
              text += `Nội dung: ${detail}\n`;
            } else {
              text += `${detail}\n`;
            }
          });
        }
        return text;
      }).join('\n');
    }
    
    return rawText;
  } catch (error: any) {
    console.error("Error extracting lessons:", error);
    if (error?.message === "QUOTA_EXCEEDED") {
      throw new Error("Bạn đã hết lượt sử dụng (Quota exceeded). Vui lòng thử lại sau vài phút hoặc đổi API Key khác.");
    }
    return rawText;
  }
};

export const generateCommentBank = async (
  subject: string,
  gradeLevel: string,
  semester: string,
  apiKey: string,
  ppct?: string,
  signal?: AbortSignal
): Promise<BankComment[]> => {
  const ai = new GoogleGenAI({ apiKey });
  
  const isTiengViet = subject === "Tiếng Việt";
  const specialSubjects = [
    "Đạo đức",
    "Tự nhiên và Xã hội",
    "Hoạt động trải nghiệm",
    "Nghệ thuật (Mỹ thuật)",
    "Nghệ thuật (Âm nhạc)",
    "Giáo dục thể chất"
  ];
  const isSpecialSubject = specialSubjects.includes(subject);

  let quantityPrompt = "";
  if (isSpecialSubject) {
    quantityPrompt = `
  YÊU CẦU SỐ LƯỢNG CHÍNH XÁC (CHO MÔN ${subject.toUpperCase()}):
  - Mức T (Hoàn thành tốt): 8 mẫu
  - Mức H (Hoàn thành): 26 mẫu
  - Mức C (Chưa hoàn thành): 5 mẫu
  Tổng cộng: 39 mẫu.
  Lưu ý: Đối với môn này, do không chấm điểm nên hãy để trường "diem" là 0 cho tất cả các mẫu. Tập trung vào nhận xét chung theo mức đạt được.`;
  } else {
    quantityPrompt = `
  YÊU CẦU SỐ LƯỢNG CHÍNH XÁC:
  - Điểm 10: 3 mẫu (Mức T - Xuất sắc)
  - Điểm 9: 3 mẫu (Mức T - Giỏi)
  - Điểm 8: 4 mẫu (Mức T - Khá giỏi)
  - Điểm 7: 6 mẫu (Mức H - Khá)
  - Điểm 6: 6 mẫu (Mức H - Trung bình khá)
  - Điểm 5: 6 mẫu (Mức H - Trung bình)
  - Điểm 4: 3 mẫu (Mức C - Yếu)
  - Điểm 3: 3 mẫu (Mức C - Kém)
  Tổng cộng: 34 mẫu.`;
  }

  const prompt = `Hãy tạo ngân hàng mẫu nhận xét cho môn ${subject}, ${gradeLevel}, học kỳ: ${semester}.
  ${isTiengViet 
    ? "ĐẶC BIỆT CHO MÔN TIẾNG VIỆT: KHÔNG ghi tên bài học cụ thể. Hãy dựa vào PPCT để nắm bắt yêu cầu kiến thức chung của học kỳ nhưng khi viết nhận xét chỉ tập trung vào mức độ đạt được của các kỹ năng Đọc, Viết, Nói, Nghe." 
    : `DỰA TRÊN DANH SÁCH BÀI HỌC TRONG PPCT SAU:
  "${ppct || "Chưa cung cấp PPCT"}"`}

  ${quantityPrompt}
  
  LƯU Ý QUAN TRỌNG: 
  1. ${isTiengViet 
      ? "TUYỆT ĐỐI KHÔNG lồng ghép tên các bài học cụ thể. Thay vào đó, hãy phân tích mức độ thành thạo các kỹ năng Đọc, Viết theo yêu cầu của học kỳ " + semester + "." 
      : isSpecialSubject 
        ? "PHẢI lồng ghép linh hoạt nội dung bài học từ PPCT vào nhận xét chung cho mức đạt được."
        : "PHẢI lồng ghép tên các bài học, bài đọc, nội dung viết từ PPCT vào nhận xét."}
  2. Nội dung nhận xét PHẢI tương xứng với từng mức đạt được/điểm cụ thể và học kỳ ${semester}.
  3. Mỗi câu nhận xét phải khác nhau, không được lặp lại.
  4. Phải bám sát Mức đạt (T, H, C) tương ứng.
  
  Yêu cầu: Nội dung mộc mạc, tiếng Việt phổ thông, không dùng từ cấm. Trả về JSON.`;

  try {
    const response = await callWithRetry(() => ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
      config: {
        systemInstruction: SYSTEM_INSTRUCTION,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              mucDo: { type: Type.STRING, enum: ["T", "H", "C"] },
              diem: { type: Type.NUMBER },
              noiDung: { type: Type.STRING },
            },
            required: ["mucDo", "diem", "noiDung"],
          },
        },
      },
    }));

    const results = JSON.parse(response.text || "[]");
    return results.map((r: any, index: number) => ({
      id: `${index + 1}`,
      ...r
    }));
  } catch (error: any) {
    console.error("Error generating bank:", error);
    if (error?.message === "QUOTA_EXCEEDED") {
      throw new Error("Bạn đã hết lượt sử dụng (Quota exceeded). Vui lòng thử lại sau vài phút hoặc đổi API Key khác.");
    }
    return [];
  }
};

export const generateComments = async (
  records: StudentRecord[], 
  subject: string, 
  gradeLevel: string,
  semester: string,
  apiKey: string,
  ppct?: string
): Promise<Partial<StudentRecord>[]> => {
  const ai = new GoogleGenAI({ apiKey });
  const isTiengViet = subject === "Tiếng Việt";
  const prompt = `Viết nhận xét (~150 ký tự) cho danh sách học sinh. 
  Môn: ${subject}. ${gradeLevel}. Học kỳ: ${semester}.
  ${isTiengViet 
    ? "ĐẶC BIỆT CHO MÔN TIẾNG VIỆT: KHÔNG được ghi tên bài học/bài đọc vào nhận xét. Hãy phân tích kỹ năng Đọc, Viết, Nói, Nghe phù hợp với mức điểm." 
    : `DỰA TRÊN DANH SÁCH BÀI HỌC TRONG PPCT SAU:
  "${ppct || "Chưa cung cấp PPCT"}"`}

  Dữ liệu học sinh: ${JSON.stringify(records.map(r => ({ stt: r.stt, mucDo: r.mucDo, diem: r.diem })))}.
  
  Ghi chú: 
  1. ${isTiengViet 
      ? "TUYỆT ĐỐI KHÔNG ghi tên bài học/bài đọc. Chỉ nhận xét kỹ năng và mức đạt được trong học kỳ " + semester + "." 
      : "PHẢI lồng ghép nội dung bài học từ PPCT vào nhận xét."}
  2. Tuyệt đối không dùng từ cấm (con, em, thầy, cô, bản, làng). 
  3. Nếu điểm bằng 0, chỉ nhận xét theo Mức đạt.`;

  try {
    const response = await callWithRetry(() => ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
      config: {
        systemInstruction: SYSTEM_INSTRUCTION,
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              stt: { type: Type.INTEGER },
              noiDung: { type: Type.STRING },
            },
            required: ["stt", "noiDung"],
          },
        },
      },
    }));
    return JSON.parse(response.text || "[]");
  } catch (error: any) {
    console.error("Error generating comments:", error);
    if (error?.message === "QUOTA_EXCEEDED") {
      throw new Error("Bạn đã hết lượt sử dụng (Quota exceeded). Vui lòng thử lại sau vài phút hoặc đổi API Key khác.");
    }
    return [];
  }
};
