
export interface StudentRecord {
  stt: number;
  maDinhDanh: string; // Mã định danh Bộ GD&ĐT
  hoTen: string;
  ngaySinh: string;
  mucDo: string; // T: Hoàn thành tốt, H: Hoàn thành, C: Chưa hoàn thành
  diem: number; // Điểm KTĐK
  maNhanXet: string;
  noiDung: string;
  thoiDiem: string; // Thời điểm đánh giá
  isProcessing?: boolean;
}

export const KHOI_LOP = [
  "Khối 1",
  "Khối 2",
  "Khối 3",
  "Khối 4",
  "Khối 5"
];

export const MON_HOC_TIEU_HOC = [
  "Tiếng Việt",
  "Toán",
  "Tiếng Anh",
  "Đạo đức",
  "Tự nhiên và Xã hội",
  "Lịch sử và Địa lý",
  "Khoa học",
  "Tin học",
  "Công nghệ",
  "Giáo dục thể chất",
  "Nghệ thuật (Âm nhạc)",
  "Nghệ thuật (Mỹ thuật)",
  "Hoạt động trải nghiệm"
];

export const getSubjectAbbr = (subject: string): string => {
  const mapping: Record<string, string> = {
    "Tiếng Việt": "TV",
    "Toán": "T",
    "Tiếng Anh": "TA",
    "Đạo đức": "DD",
    "Tự nhiên và Xã hội": "TNXH",
    "Lịch sử và Địa lý": "LSDL",
    "Khoa học": "KH",
    "Tin học": "TH",
    "Công nghệ": "CN",
    "Giáo dục thể chất": "GDTC",
    "Nghệ thuật (Âm nhạc)": "AN",
    "Nghệ thuật (Mỹ thuật)": "MT",
    "Hoạt động trải nghiệm": "HDTN"
  };
  return mapping[subject] || "MH";
};
