import React, { useState } from "react";
import {
    format,
    addDays,
    getISOWeek,
    startOfMonth,
    endOfMonth,
} from "date-fns";
import { ko } from "date-fns/locale";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import "./App.css";
import InputField from "./components/InputField";
import DailyTaskInput from "./components/InputDailTask";
import TextAreaField from "./components/InputAreaField";

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

        const firstDayOfMonth = startOfMonth(startDate);
        const monthWeekNumber =
            getISOWeek(startDate) - getISOWeek(firstDayOfMonth) + 1;

        const startWeekISO = getISOWeek(firstDayOfMonth);
        const endWeekISO = getISOWeek(endOfMonth(startDate));

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("개인 업무 일지");

        //열 넓이지정
        sheet.getColumn(1).width = 25; //A
        sheet.getColumn(2).width = 75; //B
        sheet.getColumn(3).width = 45; //C

        // A1
        sheet.getCell("A1").value = "개인작성 시트";
        sheet.getCell("A1").font = { name: "맑은 고딕", size: 11, bold: true };

        // A2:C2
        sheet.mergeCells("A2:C2");
        const a2 = sheet.getCell("A2");
        a2.value = `${
            startDate.getMonth() + 1
        }월 ${monthWeekNumber}주차 주간업무 실적 및 계획 (${startWeekISO}~${endWeekISO}주차)`;
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
        sheet.getRow(5).height = 40;
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
            aCell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };

            // 업무내용 입력칸 (3행씩)
            sheet.mergeCells(`B${rowIndex}:B${rowIndex + 2}`);
            const bCell = sheet.getCell(`B${rowIndex}`);
            bCell.value = tasks[d.toDateString()] || "";
            bCell.alignment = {
                vertical: "top",
                horizontal: "left",
                wrapText: true,
            };
            bCell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };

            rowIndex += 3;
        });

        const addSection = (row: number, title: string, value: string) => {
            // A열 병합 (제목)
            sheet.mergeCells(`A${row}:A${row + 2}`);
            const aCell = sheet.getCell(`A${row}`);
            aCell.value = title;
            aCell.font = {
                name: "맑은 고딕",
                size: 12,
                bold: true,
            };
            aCell.alignment = {
                vertical: "middle",
                horizontal: "center",
                wrapText: true,
            };
            aCell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };

            // B열 병합 (내용)
            sheet.mergeCells(`B${row}:B${row + 2}`);
            const bCell = sheet.getCell(`B${row}`);
            bCell.value = value;
            bCell.alignment = {
                vertical: "top",
                horizontal: "left",
                wrapText: true,
            };
            bCell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
            };
        };

        addSection(21, "진행 PROJECT 현황 및 ISSUE 사항", projectIssue);
        addSection(24, "개발, 개선 활동", devImprove);
        addSection(27, "출장, 연차, 휴가 계획", vacation);

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(
            new Blob([buffer]),
            `${format(new Date(), "yyyyMMdd")}-개인업무일지-${writer}.xlsx`
        );
    };

    return (
        <div
            style={{
                width: "100%",
                minHeight: "100%",
                backgroundColor: "#0d3964ff",
                display: "flex",
                justifyContent: "center",
                padding: "50px 0",
            }}
        >
            <div
                className="card p-4"
                style={{
                    maxWidth: 700,
                    width: "100%",
                    borderRadius: "15px",
                    boxShadow: "0 4px 15px rgba(0,0,0,0.3)",
                }}
            >
                <h2 className="mb-4 text-center">📑 개인 업무 일지 작성</h2>

                <InputField
                    label="날짜"
                    type="date"
                    value={date}
                    onChange={setDate}
                />
                <InputField
                    label="팀명"
                    placeholder="팀명"
                    value={team}
                    onChange={setTeam}
                />
                <InputField
                    label="작성자"
                    placeholder="작성자"
                    value={writer}
                    onChange={setWriter}
                />

                <h5 className="mt-4">💼 업무 내용</h5>
                {date &&
                    getWorkDays(new Date(date)).map((d) => (
                        <DailyTaskInput
                            key={d.toDateString()}
                            date={d}
                            value={tasks[d.toDateString()] || ""}
                            onChange={(v) =>
                                setTasks({ ...tasks, [d.toDateString()]: v })
                            }
                        />
                    ))}

                <TextAreaField
                    label="차주 계획"
                    rows={6}
                    value={nextPlan}
                    onChange={setNextPlan}
                />
                <TextAreaField
                    label="진행 PROJECT 현황 및 ISSUE 사항"
                    rows={4}
                    value={projectIssue}
                    onChange={setProjectIssue}
                />
                <TextAreaField
                    label="개발, 개선 활동"
                    rows={4}
                    value={devImprove}
                    onChange={setDevImprove}
                />
                <TextAreaField
                    label="출장, 연차, 휴가 계획"
                    rows={4}
                    value={vacation}
                    onChange={setVacation}
                />

                <button
                    className="btn btn-secondary btn-lg w-100"
                    onClick={handleExport}
                >
                    📥 엑셀 다운로드
                </button>
            </div>
        </div>
    );
};

export default App;
