
import React, { useState, useEffect, useMemo, useRef } from 'react';
import { StudentRecord, MON_HOC_TIEU_HOC, KHOI_LOP, getSubjectAbbr } from './types';
import { generateComments, generateCommentBank, BankComment, extractLessonsFromPpct } from './services/geminiService';
import { 
  Plus, 
  Sparkles, 
  Download, 
  Upload,
  BookOpen,
  CheckCircle2, 
  AlertCircle,
  Loader2,
  RefreshCw,
  Table as TableIcon,
  AlignLeft,
  Search,
  FileJson,
  Database,
  Trash2,
  FileSpreadsheet,
  Square,
  FileUp,
  FileDown,
  CalendarDays,
  Mountain,
  Settings,
  Key,
  X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import * as mammoth from 'mammoth';
import * as pdfjsLib from 'pdfjs-dist';

// Cấu hình worker cho PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

type ViewMode = 'table' | 'content';
const HOC_KY = ["Giữa kì 1", "Cuối kì 1", "Giữa kì 2", "Cuối kì 2"];

const App: React.FC = () => {
  const [records, setRecords] = useState<StudentRecord[]>([]);
  const [commentBank, setCommentBank] = useState<BankComment[]>([]);
  const [selectedSubject, setSelectedSubject] = useState(MON_HOC_TIEU_HOC[0]);
  const [selectedGrade, setSelectedGrade] = useState(KHOI_LOP[0]);
  const [selectedSemester, setSelectedSemester] = useState(HOC_KY[0]);
  const [ppct, setPpct] = useState('');
  const [rawPpctText, setRawPpctText] = useState('');
  const [ppctCache, setPpctCache] = useState<Record<string, string>>({});
  const [showPpctInput, setShowPpctInput] = useState(false);
  const [isExtractingPpct, setIsExtractingPpct] = useState(false);
  const [ppctWorkbook, setPpctWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [ppctSheetNames, setPpctSheetNames] = useState<string[]>([]);
  const [selectedPpctSheet, setSelectedPpctSheet] = useState<string>('');
  const [isGeneratingBank, setIsGeneratingBank] = useState(false);
  const [viewMode, setViewMode] = useState('table' as ViewMode);
  const [searchTerm, setSearchTerm] = useState('');
  const [notification, setNotification] = useState<{ type: 'success' | 'error' | 'info', message: string } | null>(null);
  const [apiKey, setApiKey] = useState<string>(localStorage.getItem('gemini_api_key') || '');
  const [showApiKeyModal, setShowApiKeyModal] = useState(false);
  const [tempApiKey, setTempApiKey] = useState(apiKey);
  const [studentWorkbook, setStudentWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [studentSheetNames, setStudentSheetNames] = useState<string[]>([]);
  const [selectedStudentSheet, setSelectedStudentSheet] = useState<string>('');
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const ppctFileInputRef = useRef<HTMLInputElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);

  const calculatedRecords = useMemo(() => {
    const abbr = getSubjectAbbr(selectedSubject);
    const counters: Record<string, number> = {};
    const semesterAbbr = selectedSemester === "Giữa kì 1" ? "GK1" : 
                        selectedSemester === "Cuối kì 1" ? "CK1" :
                        selectedSemester === "Giữa kì 2" ? "GK2" : "CK2";
    
    const result: StudentRecord[] = [];
    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      const studentScore = record.diem > 0 ? Math.round(record.diem) : 0;
      let level = record.mucDo;
      let thoiDiem = record.thoiDiem || semesterAbbr;
      
      if (!level && studentScore > 0) {
        if (studentScore >= 9) level = 'T';
        else if (studentScore >= 5) level = 'H';
        else level = 'C';
      }
      level = level || 'H';
      
      let targetGroup: BankComment[] = [];
      if (studentScore > 0) {
        targetGroup = commentBank.filter(b => b.diem === studentScore);
      }
      if (targetGroup.length === 0) {
        targetGroup = commentBank.filter(b => b.mucDo === level);
      }

      let autoContent = record.noiDung;
      let code = record.maNhanXet;

      if (targetGroup.length > 0) {
        // Tạo seed ngẫu nhiên nhưng ổn định dựa trên thông tin học sinh
        const nameSeed = record.hoTen.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
        // Sử dụng số nguyên tố lớn để tăng tính ngẫu nhiên và tránh mã liên tiếp
        let bankIndex = (record.stt * 137 + nameSeed * 17) % targetGroup.length;
        let selectedItem = targetGroup[bankIndex];
        
        // Kiểm tra tránh trùng với học sinh ngay trước đó nếu cùng điểm và mức đạt
        if (i > 0 && targetGroup.length > 1 && !record.maNhanXet && !record.noiDung) {
          const prev = result[i - 1];
          const prevScore = prev.diem > 0 ? Math.round(prev.diem) : 0;
          if (prevScore === studentScore && prev.mucDo === level) {
            if (selectedItem.noiDung === prev.noiDung) {
              bankIndex = (bankIndex + 1) % targetGroup.length;
              selectedItem = targetGroup[bankIndex];
            }
          }
        }
        
        if (!record.noiDung) {
          autoContent = selectedItem.noiDung;
        }
        
        if (!record.maNhanXet) {
          const itemIndexInBank = commentBank.indexOf(selectedItem);
          const sameGroupInBank = commentBank.slice(0, itemIndexInBank + 1)
            .filter(b => b.diem === selectedItem.diem && b.mucDo === selectedItem.mucDo);
          
          const itemDiemStr = selectedItem.diem || "";
          code = `${abbr}${itemDiemStr}${semesterAbbr}${selectedItem.mucDo}${sameGroupInBank.length}`;
        }
      } else if (!record.maNhanXet) {
        const counterKey = `${studentScore}_${level}`;
        counters[counterKey] = (counters[counterKey] || 0) + 1;
        const scoreStr = studentScore > 0 ? studentScore.toString() : "";
        code = `${abbr}${scoreStr}${semesterAbbr}${level}${counters[counterKey]}`;
      }

      result.push({ ...record, mucDo: level, maNhanXet: code, noiDung: autoContent, thoiDiem });
    }
    return result;
  }, [records, selectedSubject, commentBank, selectedSemester]);

  const filteredRecords = useMemo(() => {
    if (!searchTerm) return calculatedRecords;
    return calculatedRecords.filter(r => 
      r.hoTen.toLowerCase().includes(searchTerm.toLowerCase()) || 
      r.maNhanXet.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [calculatedRecords, searchTerm]);

  const handleGenerateBank = async () => {
    if (!apiKey) {
      setShowApiKeyModal(true);
      setNotification({ type: 'error', message: 'Vui lòng nhập API Key của Google AI Studio để sử dụng tính năng này.' });
      return;
    }

    if (!ppct || ppct.trim().length < 10) {
      setShowPpctInput(true);
      setNotification({ 
        type: 'error', 
        message: 'BẮT BUỘC: Vui lòng tải lên hoặc nhập Phân phối chương trình (PPCT) trước khi tạo mẫu nhận xét để AI có dữ liệu bài học.' 
      });
      // Scroll to PPCT section
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }

    setIsGeneratingBank(true);
    setNotification(null);
    abortControllerRef.current = new AbortController();
    
    try {
      const bank = await generateCommentBank(
        selectedSubject, 
        selectedGrade, 
        selectedSemester, 
        apiKey,
        ppct,
        abortControllerRef.current.signal
      );
      if (bank.length > 0) {
        setCommentBank(bank);
        setNotification({ type: 'success', message: `Đã tạo xong ngân hàng 34 mẫu nhận xét cho môn ${selectedSubject}.` });
      }
    } catch (error: any) {
      setNotification({ type: 'error', message: 'Lỗi AI khi tạo nội dung.' });
    } finally {
      setIsGeneratingBank(false);
      abortControllerRef.current = null;
    }
  };

  const handleStopGenerating = () => abortControllerRef.current?.abort();

  const parseStudentData = (wb: XLSX.WorkBook, subject: string, manualSheetName?: string) => {
    try {
      let targetSheetName = manualSheetName;
      
      if (!targetSheetName) {
        const subjectName = subject.toLowerCase();
        const subjectAbbr = getSubjectAbbr(subject).toLowerCase();
        const subPart = subject.includes('(') ? subject.split('(')[1].replace(')', '').toLowerCase().trim() : '';
        
        targetSheetName = wb.SheetNames.find(name => {
          const n = name.toLowerCase();
          return n === subjectName || n.includes(subjectName) || 
                 n === subjectAbbr || n.includes(subjectAbbr) ||
                 (subPart && n.includes(subPart));
        }) || wb.SheetNames[0];
      }

      setSelectedStudentSheet(targetSheetName);
      const ws = wb.Sheets[targetSheetName];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][];
      
      const sttKeywords = ['stt', 'số thứ tự', 'stt'];
      const idKeywords = ['mã định danh', 'mã bộ', 'mã định danh bộ', 'mã số'];
      const nameKeywords = ['họ tên', 'họ và tên', 'tên học sinh', 'học sinh', 'họ và tên học sinh'];
      const dobKeywords = ['ngày sinh', 'năm sinh', 'ns', 'ngày, tháng, năm sinh'];
      const levelKeywords = ['mức đạt được', 'mức đạt', 'xếp loại', 'đánh giá', 'mức'];
      const scoreKeywords = ['điểm ktdk', 'điểm ktđk', 'ktdk', 'điểm kiểm tra', 'điểm cuối kỳ', 'điểm'];
      const timeKeywords = ['thời điểm', 'thời điểm đánh giá', 'đánh giá tại', 'thời điểm nx'];
      
      let sttIdx = -1, idIdx = -1, nameIdx = -1, dobIdx = -1, levelIdx = -1, scoreIdx = -1, timeIdx = -1;
      let headerRowIdx = -1;

      for (let i = 0; i < Math.min(data.length, 20); i++) {
        const row = (data[i] || []).map(c => String(c).toLowerCase().trim());
        const foundName = row.findIndex(c => nameKeywords.some(k => c.includes(k)));
        
        if (foundName !== -1) {
          nameIdx = foundName;
          sttIdx = row.findIndex(c => sttKeywords.some(k => c === k || c.startsWith(k)));
          idIdx = row.findIndex(c => idKeywords.some(k => c.includes(k)));
          dobIdx = row.findIndex(c => dobKeywords.some(k => c.includes(k)));
          levelIdx = row.findIndex(c => levelKeywords.some(k => c.includes(k)));
          scoreIdx = row.findIndex(c => scoreKeywords.some(k => c.includes(k)));
          timeIdx = row.findIndex(c => timeKeywords.some(k => c.includes(k)));
          headerRowIdx = i;
          break;
        }
      }

      if (nameIdx === -1) {
        setNotification({ type: 'error', message: `Không tìm thấy cột Họ Tên trong sheet "${targetSheetName}".` });
        return;
      }

      const formatDate = (val: any): string => {
        if (!val) return "";
        if (val instanceof Date) return val.toLocaleDateString('vi-VN');
        return String(val).trim();
      };

      const newRecords = data.slice(headerRowIdx + 1)
        .filter(r => String(r[nameIdx] || "").trim().length > 1)
        .map((r, idx) => {
          let rawLevel = levelIdx !== -1 ? String(r[levelIdx] || "").trim().toUpperCase() : "";
          let level = "";
          if (rawLevel.includes("TỐT") || rawLevel === "T" || rawLevel.includes("HTT")) level = "T";
          else if (rawLevel.includes("CHƯA") || rawLevel === "C" || rawLevel.includes("CHT")) level = "C";
          else if (rawLevel.includes("HOÀN THÀNH") || rawLevel === "H" || rawLevel === "HT") level = "H";
          
          let diem = 0;
          if (scoreIdx !== -1) {
            const val = r[scoreIdx];
            if (val !== undefined && val !== "" && val !== null) {
              diem = parseFloat(String(val).replace(',', '.')) || 0;
            }
          }

          return {
            stt: sttIdx !== -1 ? (parseInt(String(r[sttIdx])) || idx + 1) : idx + 1,
            maDinhDanh: idIdx !== -1 ? String(r[idIdx] || "").trim() : "",
            hoTen: String(r[nameIdx] || "").trim(),
            ngaySinh: dobIdx !== -1 ? formatDate(r[dobIdx]) : "",
            diem: diem,
            mucDo: level,
            maNhanXet: "",
            noiDung: "",
            thoiDiem: timeIdx !== -1 ? String(r[timeIdx] || "").trim() : "",
            isProcessing: false
          } as StudentRecord;
        });

      setRecords(newRecords);
      setNotification({ 
        type: 'success', 
        message: `Đã nhập chính xác ${newRecords.length} học sinh từ sheet "${targetSheetName}".` 
      });
    } catch (err) {
      setNotification({ type: 'error', message: 'Lỗi khi xử lý dữ liệu học sinh.' });
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        setStudentWorkbook(wb);
        setStudentSheetNames(wb.SheetNames);
        parseStudentData(wb, selectedSubject);
      } catch (err) { 
        setNotification({ type: 'error', message: 'Lỗi định dạng file Excel. Hãy kiểm tra các tiêu đề cột.' }); 
      }
    };
    reader.readAsBinaryString(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Tự động chuyển sheet học sinh khi đổi môn học
  useEffect(() => {
    if (studentWorkbook) {
      // Khi đổi môn, reset sheet đã chọn để auto-detect lại cho môn mới
      setSelectedStudentSheet('');
      parseStudentData(studentWorkbook, selectedSubject);
    }
  }, [selectedSubject, studentWorkbook]);

  const handleStudentSheetChange = (sheetName: string) => {
    if (studentWorkbook) {
      setSelectedStudentSheet(sheetName);
      parseStudentData(studentWorkbook, selectedSubject, sheetName);
    }
  };

  const processPpctText = async (text: string) => {
    setRawPpctText(text);
    setPpctCache({}); // Xóa cache cũ khi tải file mới
    if (!apiKey) {
      setShowApiKeyModal(true);
      setNotification({ type: 'error', message: 'Vui lòng nhập API Key để AI trích xuất bài học.' });
      setIsExtractingPpct(false);
      return;
    }
    try {
      const cleanLessons = await extractLessonsFromPpct(text, selectedSubject, selectedGrade, apiKey);
      setPpct(cleanLessons);
      setPpctCache(prev => ({ ...prev, [selectedSubject]: cleanLessons }));
      setNotification({ 
        type: 'success', 
        message: `Đã trích xuất danh sách bài học môn ${selectedSubject} từ file.` 
      });
    } catch (err) {
      setPpct(text);
      setNotification({ type: 'error', message: 'Lỗi AI khi trích xuất bài học, đang hiển thị nội dung thô.' });
    } finally {
      setIsExtractingPpct(false);
    }
  };

  const handlePpctFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const fileName = file.name.toLowerCase();
    const reader = new FileReader();
    
    setNotification({ type: 'info', message: `Đang phân tích file ${file.name}...` });
    setIsExtractingPpct(true);
    setPpctWorkbook(null);
    setPpctSheetNames([]);
    setSelectedPpctSheet('');

    try {
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        reader.onload = async (evt) => {
          try {
            const bstr = evt.target?.result;
            const wb = XLSX.read(bstr, { type: 'binary' });
            
            setPpctWorkbook(wb);
            setPpctSheetNames(wb.SheetNames);
            
            const subjectName = selectedSubject.toLowerCase();
            const subjectAbbr = getSubjectAbbr(selectedSubject).toLowerCase();
            const subPart = selectedSubject.includes('(') ? selectedSubject.split('(')[1].replace(')', '').toLowerCase().trim() : '';
            
            const targetSheetName = wb.SheetNames.find(name => {
              const n = name.toLowerCase();
              return n === subjectName || n.includes(subjectName) || 
                     n === subjectAbbr || n.includes(subjectAbbr) ||
                     (subPart && n.includes(subPart));
            }) || wb.SheetNames[0];

            setSelectedPpctSheet(targetSheetName);
          } catch (err) {
            setNotification({ type: 'error', message: 'Lỗi khi đọc file Excel PPCT.' });
            setIsExtractingPpct(false);
          }
        };
        reader.readAsBinaryString(file);
      } 
      else if (fileName.endsWith('.docx')) {
        reader.onload = async (evt) => {
          try {
            const arrayBuffer = evt.target?.result as ArrayBuffer;
            const result = await mammoth.extractRawText({ arrayBuffer });
            await processPpctText(result.value);
          } catch (err) {
            setNotification({ type: 'error', message: 'Lỗi khi đọc file Word PPCT.' });
            setIsExtractingPpct(false);
          }
        };
        reader.readAsArrayBuffer(file);
      }
      else if (fileName.endsWith('.pdf')) {
        reader.onload = async (evt) => {
          try {
            const arrayBuffer = evt.target?.result as ArrayBuffer;
            const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
            const pdf = await loadingTask.promise;
            let fullText = "";
            
            for (let i = 1; i <= pdf.numPages; i++) {
              const page = await pdf.getPage(i);
              const textContent = await page.getTextContent();
              const pageText = textContent.items
                .map((item: any) => item.str)
                .join(" ");
              fullText += pageText + "\n";
            }
            
            await processPpctText(fullText);
          } catch (err) {
            setNotification({ type: 'error', message: 'Lỗi khi đọc file PDF PPCT.' });
            setIsExtractingPpct(false);
          }
        };
        reader.readAsArrayBuffer(file);
      }
      else {
        setNotification({ type: 'error', message: 'Định dạng file không hỗ trợ. Vui lòng chọn Excel, Word hoặc PDF.' });
        setIsExtractingPpct(false);
      }
    } catch (error) {
      setNotification({ type: 'error', message: 'Có lỗi xảy ra khi xử lý file.' });
      setIsExtractingPpct(false);
    }

    if (ppctFileInputRef.current) ppctFileInputRef.current.value = '';
  };

  // Tự động trích xuất khi đổi sheet PPCT
  useEffect(() => {
    if (ppctWorkbook && selectedPpctSheet) {
      setIsExtractingPpct(true);
      const ws = ppctWorkbook.Sheets[selectedPpctSheet];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
      
      // Lọc dữ liệu thông minh hơn: bỏ qua các cột chỉ chứa số (tuần, tiết)
      const content = data
        .map(row => {
          return row
            .filter(cell => {
              if (cell === null || cell === "") return false;
              // Bỏ qua các ô chỉ chứa số đơn lẻ (thường là tuần hoặc tiết)
              if (typeof cell === 'number') return false;
              if (typeof cell === 'string' && /^\d+$/.test(cell.trim())) return false;
              return true;
            })
            .join(" ");
        })
        .filter(text => text.trim().length > 10) // Tăng ngưỡng độ dài để loại bỏ rác
        .join("\n");
        
      processPpctText(content);
    }
  }, [selectedPpctSheet, ppctWorkbook]);

  // Tự động trích xuất lại khi đổi môn hoặc khối nếu đã có dữ liệu thô
  useEffect(() => {
    if (rawPpctText && apiKey) {
      // Kiểm tra cache trước khi gọi AI
      if (ppctCache[selectedSubject]) {
        setPpct(ppctCache[selectedSubject]);
        return;
      }

      setNotification({ type: 'info', message: `Đang trích xuất lại bài học cho môn ${selectedSubject}...` });
      setIsExtractingPpct(true);
      
      const timer = setTimeout(async () => {
        try {
          const cleanLessons = await extractLessonsFromPpct(rawPpctText, selectedSubject, selectedGrade, apiKey);
          setPpct(cleanLessons);
          setPpctCache(prev => ({ ...prev, [selectedSubject]: cleanLessons }));
          setNotification({ type: 'success', message: `Đã cập nhật bài học môn ${selectedSubject}.` });
        } catch (err) {
          setNotification({ type: 'error', message: 'Lỗi khi trích xuất lại bài học.' });
        } finally {
          setIsExtractingPpct(false);
        }
      }, 500); // Delay nhỏ để tránh spam API khi đổi nhanh
      
      return () => clearTimeout(timer);
    }
  }, [selectedSubject, selectedGrade, rawPpctText, apiKey]);

  const exportTableToExcel = () => {
    const data = filteredRecords.map(r => [
      r.stt, 
      r.maDinhDanh,
      r.hoTen, 
      r.ngaySinh,
      r.mucDo === 'T' ? 'HTT' : r.mucDo === 'H' ? 'HT' : 'CHT', 
      r.diem > 0 ? r.diem : "", 
      r.maNhanXet, 
      r.noiDung,
      r.thoiDiem
    ]);
    const ws = XLSX.utils.aoa_to_sheet([["STT", "Mã định danh Bộ GD&ĐT", "Họ và tên", "Ngày sinh", "Mức đạt được", "Điểm KTĐK", "Mã nhận xét", "Nội dung nhận xét", "Thời điểm đánh giá"], ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "NhanXet");
    XLSX.writeFile(wb, `NhanXet_${selectedGrade}_${selectedSubject}.xlsx`);
  };

  const exportBankToExcel = () => {
    if (commentBank.length === 0) {
      setNotification({ type: 'error', message: 'Chưa có ngân hàng mẫu để xuất file.' });
      return;
    }
    const abbr = getSubjectAbbr(selectedSubject);
    const semesterAbbr = selectedSemester === "Giữa kì 1" ? "GK1" : 
                        selectedSemester === "Cuối kì 1" ? "CK1" :
                        selectedSemester === "Giữa kì 2" ? "GK2" : "CK2";
    const data = commentBank.map((item, index) => {
      const sameGroup = commentBank.slice(0, index + 1).filter(b => b.diem === item.diem && b.mucDo === item.mucDo);
      const displayCode = `${abbr}${item.diem || ""}${semesterAbbr}${item.mucDo}${sameGroup.length}`;
      
      return {
        "STT": index + 1,
        "Mã nhận xét": displayCode,
        "Mức đạt": item.mucDo === 'T' ? 'HTT' : item.mucDo === 'H' ? 'HT' : 'CHT',
        "Điểm số": item.diem || "",
        "Nội dung nhận xét phổ thông": item.noiDung
      };
    });
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "NganHangMau");
    const wscols = [{wch: 5}, {wch: 15}, {wch: 12}, {wch: 10}, {wch: 95}];
    ws['!cols'] = wscols;

    XLSX.writeFile(wb, `NganHang_${selectedGrade}_${selectedSubject}.xlsx`);
    setNotification({ type: 'success', message: 'Đã xuất file ngân hàng mẫu 34 nội dung.' });
  };

  const handleSaveApiKey = () => {
    localStorage.setItem('gemini_api_key', tempApiKey);
    setApiKey(tempApiKey);
    setShowApiKeyModal(false);
    setNotification({ type: 'success', message: 'Đã lưu API Key thành công!' });
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] pb-20 text-slate-900">
      {/* API Key Modal */}
      {showApiKeyModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-[2rem] shadow-2xl w-full max-w-md overflow-hidden border border-slate-200 animate-in zoom-in-95 duration-200">
            <div className="p-8">
              <div className="flex items-center justify-between mb-6">
                <div className="flex items-center gap-3">
                  <div className="p-2.5 bg-indigo-100 text-indigo-600 rounded-xl">
                    <Key size={20} />
                  </div>
                  <h3 className="text-xl font-black uppercase tracking-tight text-slate-800">Cấu hình API Key</h3>
                </div>
                <button onClick={() => setShowApiKeyModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                  <X size={24} />
                </button>
              </div>
              
              <p className="text-sm text-slate-500 mb-6 font-medium leading-relaxed">
                Hãy nhập API Key từ <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener noreferrer" className="text-indigo-600 font-bold hover:underline">Google AI Studio</a> để sử dụng các tính năng thông minh của trợ lý.
              </p>
              
              <div className="space-y-4">
                <div>
                  <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 ml-1">Google Gemini API Key</label>
                  <input 
                    type="password" 
                    placeholder="Dán API Key của bạn vào đây..." 
                    value={tempApiKey}
                    onChange={(e) => setTempApiKey(e.target.value)}
                    className="w-full px-5 py-4 bg-slate-50 border border-slate-200 rounded-2xl text-sm outline-none focus:ring-4 focus:ring-indigo-100 focus:border-indigo-300 transition-all font-mono"
                  />
                </div>
                
                <div className="flex gap-3 pt-2">
                  <button 
                    onClick={() => setShowApiKeyModal(false)}
                    className="flex-1 px-6 py-4 bg-slate-100 text-slate-600 rounded-2xl text-sm font-bold hover:bg-slate-200 transition-all active:scale-95"
                  >
                    Hủy bỏ
                  </button>
                  <button 
                    onClick={handleSaveApiKey}
                    className="flex-1 px-6 py-4 bg-indigo-600 text-white rounded-2xl text-sm font-bold shadow-lg shadow-indigo-200 hover:bg-indigo-700 transition-all active:scale-95"
                  >
                    Lưu cấu hình
                  </button>
                </div>
              </div>
            </div>
            <div className="bg-slate-50 px-8 py-4 border-t border-slate-100">
              <p className="text-[10px] text-slate-400 font-bold text-center uppercase tracking-tighter">
                API Key được lưu an toàn trong trình duyệt của bạn
              </p>
            </div>
          </div>
        </div>
      )}

      <header className="bg-white border-b sticky top-0 z-40 px-4 md:px-8 py-4 shadow-sm">
        <div className="w-full flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-indigo-600 rounded-xl text-white shadow-lg relative">
              <BookOpen size={24} />
              <Mountain size={14} className="absolute -bottom-1 -right-1 text-white opacity-40" />
            </div>
            <div>
              <h1 className="text-xl font-black uppercase tracking-tight">TRỢ LÍ TẠO NHẬN XÉT  ({selectedGrade})</h1>
              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest leading-tight">Môn: {selectedSubject}  </p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <div className="flex bg-slate-100 p-1 rounded-xl items-center border border-slate-200">
              <select value={selectedGrade} onChange={(e) => setSelectedGrade(e.target.value)} className="px-3 py-1.5 bg-transparent text-sm font-bold outline-none border-r border-slate-200 cursor-pointer">
                {KHOI_LOP.map(k => <option key={k} value={k}>{k}</option>)}
              </select>
              <select value={selectedSubject} onChange={(e) => setSelectedSubject(e.target.value)} className="px-3 py-1.5 bg-transparent text-sm font-bold outline-none border-r border-slate-200 cursor-pointer">
                {MON_HOC_TIEU_HOC.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
              <select value={selectedSemester} onChange={(e) => setSelectedSemester(e.target.value)} className="px-3 py-1.5 bg-transparent text-sm font-bold outline-none text-indigo-600 cursor-pointer">
                {HOC_KY.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <button 
              onClick={() => setShowPpctInput(!showPpctInput)} 
              className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold transition-all border relative ${showPpctInput ? 'bg-indigo-50 border-indigo-200 text-indigo-600' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}
            >
              <AlignLeft size={16} /> Phân phối chương trình 
              {!ppct && (
                <span className="absolute -top-1 -right-1 flex h-3 w-3">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-3 w-3 bg-rose-500"></span>
                </span>
              )}
            </button>

            <div className="h-8 w-px bg-slate-200 mx-1 hidden sm:block"></div>

            <button 
              onClick={() => setShowApiKeyModal(true)}
              className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold transition-all border ${apiKey ? 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50' : 'bg-rose-50 border-rose-200 text-rose-600 animate-pulse'}`}
              title="Cấu hình API Key"
            >
              <Settings size={16} /> {apiKey ? 'Cài đặt' : 'Nhập API'}
            </button>

            {isGeneratingBank ? (
              <button onClick={handleStopGenerating} className="inline-flex items-center gap-2 px-5 py-2.5 bg-rose-600 text-white rounded-xl shadow-md text-sm font-bold animate-pulse">
                <Square size={16} fill="currentColor" /> Dừng tạo
              </button>
            ) : (
              <button 
                onClick={handleGenerateBank} 
                className={`inline-flex items-center gap-2 px-5 py-2.5 rounded-xl shadow-md text-sm font-bold transition-all group active:scale-95 relative ${!ppct ? 'bg-slate-400 text-white' : 'bg-amber-500 text-white hover:bg-amber-600'}`}
              >
                <Database size={16} className="group-hover:rotate-12 transition-transform" /> 
                Tạo mẫu nhận xét
                {!ppct && (
                  <span className="absolute -top-2 -right-2 flex h-4 w-4">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-4 w-4 bg-rose-500 items-center justify-center text-[8px] text-white font-black">!</span>
                  </span>
                )}
              </button>
            )}

            <input type="file" accept=".xlsx,.xls" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
            {studentSheetNames.length > 0 ? (
              <div className="flex items-center bg-slate-100 p-1 rounded-xl border border-slate-200">
                <button 
                  onClick={() => fileInputRef.current?.click()} 
                  className="p-2 text-slate-500 hover:text-indigo-600 transition-colors" 
                  title="Tải file CSDL khác"
                >
                  <RefreshCw size={16} />
                </button>
                <div className="w-px h-4 bg-slate-300 mx-1"></div>
                <select 
                  value={selectedStudentSheet} 
                  onChange={(e) => handleStudentSheetChange(e.target.value)}
                  className="px-3 py-1.5 bg-transparent text-sm font-bold outline-none text-slate-700 cursor-pointer max-w-[180px] truncate"
                >
                  {studentSheetNames.map(name => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>
            ) : (
              <button onClick={() => fileInputRef.current?.click()} className="inline-flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-600 rounded-xl text-sm font-bold hover:bg-slate-50 transition-colors">
                <FileUp size={16} /> Nhập file đánh giá CSDL 
              </button>
            )}

            {viewMode === 'table' ? (
              <button onClick={exportTableToExcel} className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-xl shadow-md text-sm font-bold active:scale-95 transition-all">
                <Download size={16} /> Xuất nhận xét
              </button>
            ) : (
              <button onClick={exportBankToExcel} className="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-xl shadow-md text-sm font-bold active:scale-95 transition-all">
                <FileSpreadsheet size={16} /> Xuất ngân hàng (.xlsx)
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="w-full px-4 md:px-8 lg:px-12 mt-8">
        {showPpctInput && (
          <div className="mb-8 bg-white p-6 rounded-[2rem] shadow-xl border border-indigo-100 animate-in slide-in-from-top duration-300">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <div className="p-2 bg-indigo-100 text-indigo-600 rounded-lg">
                  <AlignLeft size={18} />
                </div>
                <h3 className="font-black uppercase tracking-tight text-slate-700">Phân phối chương trình (PPCT)</h3>
              </div>
              <button onClick={() => setShowPpctInput(false)} className="text-slate-400 hover:text-slate-600">
                <Plus size={20} className="rotate-45" />
              </button>
            </div>
            <p className="text-xs text-slate-500 mb-4 font-medium italic">
              * Tải lên file Phân phối chương trình ( Excel, PDF) để AI phân tích các bài học và chủ đề đã học trong học kỳ này.
            </p>
            
            <div className="flex flex-col gap-4">
              <div className="flex items-center gap-3">
                <input 
                  type="file" 
                  accept=".xlsx,.xls,.docx,.pdf" 
                  className="hidden" 
                  ref={ppctFileInputRef} 
                  onChange={handlePpctFileUpload} 
                />
                <button 
                  onClick={() => ppctFileInputRef.current?.click()}
                  disabled={isExtractingPpct}
                  className={`inline-flex items-center gap-2 px-6 py-3 text-white rounded-2xl shadow-lg text-sm font-bold transition-all active:scale-95 ${isExtractingPpct ? 'bg-slate-400 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700'}`}
                >
                  {isExtractingPpct ? (
                    <>
                      <Loader2 size={18} className="animate-spin" /> Đang trích xuất bài học...
                    </>
                  ) : (
                    <>
                      <Upload size={18} /> Tải lên file PPCT ( Excel, PDF)
                    </>
                  )}
                </button>
                {ppct && !isExtractingPpct && (
                  <button 
                    onClick={() => {
                      setPpct('');
                      setPpctWorkbook(null);
                      setPpctSheetNames([]);
                      setSelectedPpctSheet('');
                    }}
                    className="text-xs text-rose-500 font-bold hover:underline"
                  >
                    Xóa dữ liệu đã tải
                  </button>
                )}
              </div>

              {ppctSheetNames.length > 0 && (
                <div className="flex items-center gap-3 bg-indigo-50 p-3 rounded-xl border border-indigo-100">
                  <label className="text-xs font-black uppercase tracking-wider text-indigo-600 whitespace-nowrap">Chọn Sheet PPCT:</label>
                  <select 
                    value={selectedPpctSheet} 
                    onChange={(e) => setSelectedPpctSheet(e.target.value)}
                    className="flex-1 px-3 py-2 bg-white border border-indigo-200 rounded-lg text-sm font-bold outline-none focus:ring-2 focus:ring-indigo-500 transition-all"
                  >
                    {ppctSheetNames.map(name => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                </div>
              )}

              {isExtractingPpct ? (
                <div className="border-2 border-dashed border-indigo-200 rounded-2xl p-12 text-center bg-indigo-50/30">
                  <Loader2 size={40} className="mx-auto mb-4 text-indigo-400 animate-spin" />
                  <p className="text-sm text-indigo-600 font-black uppercase tracking-widest">AI đang đọc và liệt kê danh sách bài học...</p>
                  <p className="text-xs text-slate-400 mt-2 font-medium italic">Vui lòng chờ trong giây lát, quá trình này có thể mất vài giây.</p>
                </div>
              ) : ppct ? (
                <div className="bg-slate-50 border border-slate-200 rounded-2xl p-6">
                  <div className="flex items-center justify-between mb-4">
                    <div className="flex items-center gap-2 text-emerald-600">
                      <CheckCircle2 size={16} />
                      <span className="text-xs font-black uppercase tracking-wider">Danh sách bài học đã trích xuất</span>
                    </div>
                    <span className="text-[10px] text-slate-400 font-bold">{ppct.split('\n').filter(l => l.trim()).length} bài học/chủ đề</span>
                  </div>
                  <div className="text-sm text-slate-600 font-medium leading-relaxed max-h-[500px] overflow-y-auto pr-4 custom-scrollbar bg-white p-6 rounded-xl border border-slate-100 shadow-inner">
                    {ppct.split('\n').map((line, i) => {
                      const trimmedLine = line.trim();
                      const isTitle = trimmedLine.toUpperCase().startsWith('BÀI') || trimmedLine.toUpperCase().startsWith('CHỦ ĐỀ');
                      
                      // Check for labels (any word ending with a colon)
                      let styledLine: React.ReactNode = line;
                      const colonIndex = trimmedLine.indexOf(':');
                      
                      if (!isTitle && colonIndex > 0 && colonIndex < 30) {
                        const label = trimmedLine.substring(0, colonIndex + 1);
                        styledLine = (
                          <>
                            <span className="font-black text-slate-800">{label}</span>
                            {line.substring(line.indexOf(label) + label.length)}
                          </>
                        );
                      }

                      return (
                        <div 
                          key={i} 
                          className={`${isTitle ? 'text-rose-600 font-black mt-4 mb-1' : 'ml-4'} ${trimmedLine === '' ? 'h-2' : ''}`}
                        >
                          {styledLine}
                        </div>
                      );
                    })}
                  </div>

                  <div className="mt-8 flex justify-center">
                    <button 
                      onClick={() => {
                        setViewMode('content');
                        handleGenerateBank();
                      }}
                      disabled={isGeneratingBank}
                      className="inline-flex items-center gap-3 px-10 py-5 bg-amber-500 text-white rounded-[1.5rem] shadow-2xl text-lg font-black hover:bg-amber-600 transition-all active:scale-95 group hover:shadow-amber-200/50"
                    >
                      {isGeneratingBank ? (
                        <>
                          <Loader2 size={24} className="animate-spin" /> Đang tạo ngân hàng mẫu...
                        </>
                      ) : (
                        <>
                          <Sparkles size={24} className="group-hover:animate-bounce" /> 
                          Tạo ngân hàng mẫu từ dữ liệu này
                        </>
                      )}
                    </button>
                  </div>
                </div>
              ) : (
                <div className="border-2 border-dashed border-slate-200 rounded-2xl p-8 text-center">
                  <FileSpreadsheet size={32} className="mx-auto mb-2 text-slate-300" />
                  <p className="text-xs text-slate-400 font-bold uppercase tracking-tight">Chưa có file PPCT nào được chọn</p>
                </div>
              )}
            </div>
          </div>
        )}

        <div className="flex flex-col sm:flex-row justify-between items-center mb-8 gap-4">
          <div className="bg-white p-1.5 rounded-2xl shadow-sm border border-slate-200 flex">
            <button onClick={() => setViewMode('table')} className={`px-8 py-2.5 rounded-xl text-sm font-bold transition-all ${viewMode === 'table' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-500 hover:bg-slate-50'}`}>
              <TableIcon size={18} className="inline mr-2" /> Học sinh
            </button>
            <button onClick={() => setViewMode('content')} className={`px-8 py-2.5 rounded-xl text-sm font-bold transition-all ${viewMode === 'content' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-500 hover:bg-slate-50'}`}>
              <FileJson size={18} className="inline mr-2" /> Ngân hàng mẫu
            </button>
          </div>
          <div className="relative w-full sm:w-96">
            <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <input type="text" placeholder="Tìm tên, ngày sinh, mã..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-12 pr-4 py-3 bg-white border border-slate-300 rounded-2xl text-sm outline-none shadow-sm focus:ring-4 focus:ring-indigo-100 transition-all font-medium" />
          </div>
        </div>

        {notification && (
          <div className={`mb-6 p-4 rounded-2xl flex items-center justify-between border shadow-sm ${notification.type === 'success' ? 'bg-emerald-50 text-emerald-900 border-emerald-200' : 'bg-rose-50 text-rose-900 border-rose-200'}`}>
            <div className="flex items-center gap-3">
              <CheckCircle2 size={20} className={notification.type === 'success' ? 'text-emerald-600' : 'text-rose-600'} />
              <span className="font-bold text-sm tracking-tight">{notification.message}</span>
            </div>
            <button onClick={() => setNotification(null)} className="opacity-50 hover:opacity-100 text-xs uppercase font-black">Đóng</button>
          </div>
        )}

        <div className="bg-white rounded-[2rem] shadow-2xl border border-slate-200 overflow-hidden mb-12">
          {viewMode === 'table' ? (
            <div className="overflow-x-auto">
              <table className="table-bordered min-w-[1100px] border-none text-center">
                <thead className="bg-slate-50 text-slate-600 font-bold border-b-2 border-slate-200">
                  <tr>
                    <th className="px-4 py-4 w-12 text-slate-400">STT</th>
                    <th className="px-4 py-4 w-40">Mã định danh Bộ GD&ĐT</th>
                    <th className="px-6 py-4 w-56 text-left">Họ và tên</th>
                    <th className="px-4 py-4 w-32">Ngày sinh</th>
                    <th className="px-4 py-4 w-28">Mức đạt được</th>
                    <th className="px-4 py-4 w-24">Điểm KTĐK</th>
                    <th className="px-4 py-4 w-32">Mã nhận xét</th>
                    <th className="px-6 py-4 text-left">Nội dung nhận xét</th>
                    <th className="px-4 py-4 w-32">Thời điểm đánh giá</th>
                    <th className="px-4 py-4 w-16 text-slate-400">Xóa</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredRecords.map((r) => (
                    <tr key={r.stt} className="group hover:bg-indigo-50/40 transition-all">
                      <td className="px-4 py-5 text-xs text-slate-400 font-bold">{r.stt}</td>
                      <td className="px-4 py-5 text-xs text-slate-600 font-medium">{r.maDinhDanh}</td>
                      <td className="px-6 py-5 text-sm font-bold text-slate-900 text-left">{r.hoTen}</td>
                      <td className="px-4 py-5 text-xs font-medium text-slate-500">{r.ngaySinh || "-"}</td>
                      <td className="px-4 py-5">
                        <span className={`px-3 py-1 rounded-lg text-[9px] font-black border uppercase shadow-sm ${r.mucDo === 'T' ? 'bg-amber-50 text-amber-700 border-amber-200' : r.mucDo === 'H' ? 'bg-sky-50 text-sky-700 border-sky-200' : r.mucDo === 'C' ? 'bg-rose-50 text-rose-700 border-rose-200' : 'bg-slate-50'}`}>
                          {r.mucDo === 'T' ? 'HTT' : r.mucDo === 'H' ? 'HT' : r.mucDo === 'C' ? 'CHT' : 'TRỐNG'}
                        </span>
                      </td>
                      <td className="px-4 py-5 font-black text-slate-700 text-lg">{r.diem || "-"}</td>
                      <td className="px-4 py-5">
                        <span className="text-[10px] font-mono font-black bg-white px-2 py-1.5 rounded-lg border border-slate-200 text-indigo-600 shadow-sm">{r.maNhanXet}</span>
                      </td>
                      <td className="px-6 py-5">
                        <textarea value={r.noiDung} onChange={(e) => setRecords(records.map(rec => rec.stt === r.stt ? { ...rec, noiDung: e.target.value } : rec))} className="w-full bg-transparent border-none text-sm font-medium resize-none min-h-[60px] outline-none leading-relaxed focus:bg-white focus:p-2 focus:rounded-lg transition-all" placeholder="Thông tin STT, Họ tên, NS, Điểm, Mức sẽ tự hiển thị sau khi nhập Excel..." />
                      </td>
                      <td className="px-4 py-5 text-xs text-slate-600 font-bold">{r.thoiDiem}</td>
                      <td className="px-4 py-5">
                        <button onClick={() => setRecords(records.filter(rec => rec.stt !== r.stt))} className="text-slate-200 hover:text-rose-500 opacity-0 group-hover:opacity-100 transition-all"><Trash2 size={16} /></button>
                      </td>
                    </tr>
                  ))}
                  {filteredRecords.length === 0 && (
                    <tr>
                      <td colSpan={10} className="py-40 text-center opacity-30">
                        <Mountain size={64} className="mx-auto mb-4" />
                        <p className="font-bold uppercase tracking-widest text-sm">Chưa có dữ liệu học sinh. Hãy Nhập Excel môn {selectedSubject}.</p>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="table-bordered w-full text-left border-none text-center">
                <thead className="bg-slate-50 text-slate-500 font-bold uppercase text-[10px] tracking-widest border-b border-slate-200">
                  <tr>
                    <th className="px-8 py-6 w-56">Mã nhận xét</th>
                    <th className="px-8 py-6 w-40">Mức đạt</th>
                    <th className="px-8 py-6 w-24">Điểm</th>
                    <th className="px-8 py-6 text-left">Mẫu nhận xét phổ thông</th>
                    <th className="px-4 py-6 w-16 text-slate-300">Xóa</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {commentBank.map((item, idx) => {
                    const abbr = getSubjectAbbr(selectedSubject);
                    const semesterAbbr = selectedSemester === "Giữa kì 1" ? "GK1" : 
                                        selectedSemester === "Cuối kì 1" ? "CK1" :
                                        selectedSemester === "Giữa kì 2" ? "GK2" : "CK2";
                    const sameGroup = commentBank.slice(0, idx + 1).filter(b => b.diem === item.diem && b.mucDo === item.mucDo);
                    const displayCode = `${abbr}${item.diem || ""}${semesterAbbr}${item.mucDo}${sameGroup.length}`;
                    return (
                      <tr key={item.id} className="hover:bg-indigo-50/20 transition-all">
                        <td className="px-8 py-8">
                          <span className="text-xs font-black font-mono px-4 py-2 bg-indigo-50 text-indigo-700 border border-indigo-100 rounded-xl shadow-sm">{displayCode}</span>
                        </td>
                        <td className="px-8 py-8">
                          <span className={`px-4 py-2 rounded-xl text-[10px] font-black border uppercase inline-block min-w-[140px] shadow-sm ${item.mucDo === 'T' ? 'bg-amber-50 text-amber-700 border-amber-200' : item.mucDo === 'H' ? 'bg-sky-50 text-sky-700 border-sky-200' : 'bg-rose-50 text-rose-700 border-rose-200'}`}>
                            {item.mucDo === 'T' ? 'Tốt (10,9,8)' : item.mucDo === 'H' ? 'HT (7,6,5)' : 'CHT (4,3)'}
                          </span>
                        </td>
                        <td className="px-8 py-8 font-black text-slate-700 text-xl">{item.diem || "-"}</td>
                        <td className="px-8 py-8">
                          <textarea value={item.noiDung} onChange={(e) => setCommentBank(commentBank.map(b => b.id === item.id ? { ...b, noiDung: e.target.value } : b))} className="w-full bg-white border border-slate-200 rounded-2xl p-6 text-sm font-medium leading-relaxed min-h-[80px] outline-none shadow-sm focus:ring-4 focus:ring-indigo-50 transition-all resize-none" />
                        </td>
                        <td className="px-4 py-8">
                          <button onClick={() => setCommentBank(commentBank.filter(b => b.id !== item.id))} className="text-slate-200 hover:text-rose-500 transition-all"><Trash2 size={18} /></button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>

      <footer className="mt-20 border-t border-slate-200 py-16 text-center bg-white/50">
        <div className="flex items-center justify-center gap-2 text-slate-300 mb-2">
          <Mountain size={16} />
          <span className="text-[10px] font-black uppercase tracking-[0.3em]">Hỗ trợ giáo viên tiểu học</span>
        </div>
        <div className="text-[11px] text-slate-400 font-bold uppercase tracking-widest">
          TRƯỜNG PTDTBT TH Nấm Dẩn | Thầy Nguyễn Đức Hùng | Phiên bản 12.0 - Tự động ghi nhận thông tin chính xác
        </div>
      </footer>
    </div>
  );
};

export default App;
