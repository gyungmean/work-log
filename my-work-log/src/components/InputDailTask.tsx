// DailyTaskInput.tsx
import React from "react";
import { format } from "date-fns";
import { ko } from "date-fns/locale";

interface Props {
    date: Date;
    value: string;
    onChange: (value: string) => void;
}

const DailyTaskInput: React.FC<Props> = ({ date, value, onChange }) => (
    <div className="mb-3">
        <label className="form-label">
            {format(date, "MM/dd (EEE)", { locale: ko })}
        </label>
        <textarea
            className="form-control"
            style={{ height: 60 }}
            value={value}
            onChange={(e) => onChange(e.target.value)}
        />
    </div>
);

export default DailyTaskInput;
