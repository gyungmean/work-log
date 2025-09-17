import React, { useState } from "react";
import { format, addDays, getWeek, getMonth } from "date-fns";
import { ko } from "date-fns/locale";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const App: React.FC = () => {
    const [date, setDate] = useState<string>("");
    const [team, setTeam] = useState("");
    const [writer, setWriter] = useState("");
    const [tasks, setTasks] = useState<{ [key: string]: string }>({});
    const [nextPlan, setNextPlan] = useState("");
    const [projectIssue, setProjectIssue] = useState("");
    const [devImprove, setDevImprove] = useState("");
    const [vacation, setVacation] = useState("");

    const getWorkDays = (startDate: Date) => {
        const days: Date[] = [];
        let d = new Date(startDate);
        while (days.length < 5) {
            const day = d.getDay();
            if (day !== 0 && day !== 6) {
                days.push(new Date(d));
            }
            d = addDays(d, 1);
        }
        return days;
    };

    const handleExport = async () => {
        if (!date) return alert("날짜를 선택해주세요!");

        const startDate = new Date(date);
        const workDays = getWorkDays(startDate);

        const month = getMonth(startDate) + 1;
        const weekNumber = getWeek(startDate, { locale: ko });
        const startWeek = getWeek(workDays[0], { locale: ko });
        const endWeek = getWeek(workDays[4], { locale: ko });

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("개인 업무 일지");

        // A1
        sheet.getCell("A1").value = "개인작성 시트";
        sheet.getCell("A1").font = { name: "맑은 고딕", size: 11, bold: true };

        // A2:C2
        sheet.mergeCells("A2:C2");
        const a2 = sheet.getCell("A2");
        a2.value = `${month}월 ${weekNumber}주차 주간업무 실적 및 계획 (${startWeek}-${endWeek}주차)`;
        a2.font = { name: "맑은 고딕", size: 24, underline: true };
        a2.alignment = { horizontal: "center", vertical: "middle" };

        // A3
        sheet.addRow([]);

        // A4 / C4
        sheet.getCell("A4").value = `- ${team}팀`;
        sheet.getCell("A4").font = { name: "맑은 고딕", size: 14 };

        sheet.getCell("C4").value = `작성자 : ◆ ${writer} (${format(
            new Date(),
            "yyyy-MM-dd"
        )})`;
        sheet.getCell("C4").font = { name: "맑은 고딕", size: 14 };

        // A5:C5
        sheet.getRow(5).height = 25;
        sheet.getCell("B5").value = `금주 실적 (${format(
            workDays[0],
            "MM월 dd일"
        )} ~ ${format(workDays[4], "MM월 dd일")})`;
        sheet.getCell("C5").value = `차주 계획`;
        ["A5", "B5", "C5"].forEach((key) => {
            const cell = sheet.getCell(key);
            cell.font = { name: "맑은 고딕", size: 14, bold: true };
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "D9D9D9" },
            };
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };
        });

        // C열 병합 (차주 계획)
        sheet.mergeCells("C6:C29");
        sheet.getCell("C6").value = nextPlan;
        sheet.getCell("C6").alignment = {
            vertical: "top",
            horizontal: "left",
            wrapText: true,
        };

        // 날짜 및 업무내용
        let rowIndex = 6;
        workDays.forEach((d) => {
            // 날짜 병합 (3행씩)
            sheet.mergeCells(`A${rowIndex}:A${rowIndex + 2}`);
            const aCell = sheet.getCell(`A${rowIndex}`);
            aCell.value = format(d, "MM/dd (EEE)", { locale: ko });
            aCell.alignment = { vertical: "middle", horizontal: "center" };

            // 업무내용 입력칸 (3행씩)
            sheet.mergeCells(`B${rowIndex}:B${rowIndex + 2}`);
            const bCell = sheet.getCell(`B${rowIndex}`);
            bCell.value = tasks[d.toDateString()] || "";
            bCell.alignment = {
                vertical: "top",
                horizontal: "left",
                wrapText: true,
            };

            rowIndex += 3;
        });

        // 📌 추가 섹션
        const addSection = (row: number, title: string, value: string) => {
            sheet.getCell(`A${row}`).value = title;
            sheet.getCell(`A${row}`).font = {
                name: "맑은 고딕",
                size: 12,
                bold: true,
            };
            sheet.mergeCells(`B${row}:B${row + 2}`);
            sheet.getCell(`B${row}`).value = value;
            sheet.getCell(`B${row}`).alignment = {
                vertical: "top",
                horizontal: "left",
                wrapText: true,
            };
        };

        addSection(21, "진행 PROJECT 현황 및 ISSUE 사항", projectIssue);
        addSection(24, "개발, 개선 활동", devImprove);
        addSection(27, "출장, 연차, 휴가 계획", vacation);

        // 전체 테두리
        sheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: "thin" },
                    left: { style: "thin" },
                    bottom: { style: "thin" },
                    right: { style: "thin" },
                };
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(
            new Blob([buffer]),
            `${format(new Date(), "yyyyMMdd")}-개인업무일지-${writer}.xlsx`
        );
    };

    return (
        <div style={{ padding: 20 }}>
            <h1>📑 개인 업무 일지 작성</h1>
            <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
            />
            <input
                placeholder="팀명"
                value={team}
                onChange={(e) => setTeam(e.target.value)}
            />
            <input
                placeholder="작성자"
                value={writer}
                onChange={(e) => setWriter(e.target.value)}
            />

            <h3>업무 내용</h3>
            {date &&
                getWorkDays(new Date(date)).map((d) => (
                    <div key={d.toDateString()}>
                        {format(d, "MM/dd (EEE)", { locale: ko })} :
                        <textarea
                            style={{ width: 400, height: 60 }}
                            value={tasks[d.toDateString()] || ""}
                            onChange={(e) =>
                                setTasks({
                                    ...tasks,
                                    [d.toDateString()]: e.target.value,
                                })
                            }
                        />
                    </div>
                ))}

            <h3>차주 계획</h3>
            <textarea
                style={{ width: 400, height: 120 }}
                value={nextPlan}
                onChange={(e) => setNextPlan(e.target.value)}
            />

            <h3>진행 PROJECT 현황 및 ISSUE 사항</h3>
            <textarea
                style={{ width: 400, height: 80 }}
                value={projectIssue}
                onChange={(e) => setProjectIssue(e.target.value)}
            />

            <h3>개발, 개선 활동</h3>
            <textarea
                style={{ width: 400, height: 80 }}
                value={devImprove}
                onChange={(e) => setDevImprove(e.target.value)}
            />

            <h3>출장, 연차, 휴가 계획</h3>
            <textarea
                style={{ width: 400, height: 80 }}
                value={vacation}
                onChange={(e) => setVacation(e.target.value)}
            />

            <br />
            <button onClick={handleExport}>📥 엑셀 다운로드</button>
        </div>
    );
};

export default App;
