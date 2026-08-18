const AmLich = (function() {
    const CAN = ["Giáp", "Ất", "Bính", "Đinh", "Mậu", "Kỷ", "Canh", "Tân", "Nhâm", "Quý"];
    const CHI = ["Tý", "Sửu", "Dần", "Mão", "Thìn", "Tỵ", "Ngọ", "Mùi", "Thân", "Dậu", "Tuất", "Hợi"];

    /**
     * LUNAR_DATA lưu trữ thông tin nén của từng năm:
     * - data: Mã Hex 16-bit (Bit 1: 30 ngày, Bit 0: 29 ngày).
     * - leap: Tháng nhuận (0 nếu không nhuận).
     * - tet: [Ngày, Tháng] Dương lịch của ngày Mùng 1 Tết.
     */
    const LUNAR_DATA = {
        2024: { data: 0x0a930, leap: 0, tet: [10, 2] },
        2025: { data: 0x052b0, leap: 6, tet: [29, 1] }, 
        2026: { data: 0x0a570, leap: 0, tet: [17, 2] }, // Tết Bính Ngọ 17/02/2026
        2027: { data: 0x04ad0, leap: 0, tet: [6, 2] },
        2028: { data: 0x05a6d, leap: 5, tet: [26, 1] },
        2029: { data: 0x0a6d0, leap: 0, tet: [13, 2] },
        2030: { data: 0x052d0, leap: 0, tet: [3, 2] }
    };

    function getJulianDay(d, m, y) {
        let a = Math.floor((14 - m) / 12);
        y = y + 4800 - a;
        m = m + 12 * a - 3;
        return d + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) - 32045;
    }

    function getCanChiThang(lMonth, lYear) {
        let canYear = (lYear + 6) % 10;
        let canIdx = (canYear * 2 + 2 + (lMonth - 1)) % 10;
        return CAN[canIdx] + " " + CHI[(lMonth + 1) % 12];
    }

    return {        
        convert: function(d, m, y) {
            const jd = getJulianDay(d, m, y);
            let info = LUNAR_DATA[y];
            if (!info) return null;

            let jdTet = getJulianDay(info.tet[0], info.tet[1], y);
            let lunarYear = y;
            let diff = jd - jdTet;

            // Nếu ngày nằm trước Tết, lấy dữ liệu năm trước
            if (diff < 0) {
                lunarYear = y - 1;
                info = LUNAR_DATA[lunarYear];
                if (!info) return null;
                jdTet = getJulianDay(info.tet[0], info.tet[1], lunarYear);
                diff = jd - jdTet;
            }

            let lunarMonth = 1, lunarDay = 1, isLeapMonth = false, tempDiff = diff;
            const totalMonths = info.leap > 0 ? 13 : 12;

            // SỬA TẠI ĐÂY: Thuật toán quét bit chính xác cho mã Hex 16-bit
            for (let i = 0; i < totalMonths; i++) {
                // Lấy bit thứ (15 - i) từ mã Hex. Bit 1 = 30 ngày, Bit 0 = 29 ngày
                let daysInMonth = ((info.data >> (15 - i)) & 1) === 1 ? 30 : 29;
                
                if (tempDiff < daysInMonth) {
                    lunarDay = tempDiff + 1;
                    if (info.leap > 0) {
                        if (i + 1 <= info.leap) lunarMonth = i + 1;
                        else if (i + 1 === info.leap + 1) { 
                            lunarMonth = info.leap; 
                            isLeapMonth = true; 
                        } else { 
                            lunarMonth = i; 
                        }
                    } else { 
                        lunarMonth = i + 1; 
                    }
                    break;
                }
                tempDiff -= daysInMonth;
            }

            return {
                day: lunarDay,
                month: lunarMonth,
                year: lunarYear,
                isLeap: isLeapMonth,
                canChiNgay: CAN[(jd + 9) % 10] + " " + CHI[(jd + 1) % 12],
                canChiThang: getCanChiThang(lunarMonth, lunarYear),
                canChiNam: CAN[(lunarYear + 6) % 10] + " " + CHI[(lunarYear + 8) % 12],
                fullDate: `${lunarDay}/${lunarMonth}${isLeapMonth ? ' (Nhuận)' : ''}/${lunarYear}`
            };
        }
    };
})();