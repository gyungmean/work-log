// TextAreaField.tsx
import React from "react";

interface Props {
    label: string;
    value: string;
    rows?: number;
    onChange: (value: string) => void;
}

const TextAreaField: React.FC<Props> = ({
    label,
    value,
    rows = 3,
    onChange,
}) => (
    <div className="mb-3">
        <label className="form-label">{label}</label>
        <textarea
            className="form-control"
            style={{ height: rows * 10 }}
            value={value}
            onChange={(e) => onChange(e.target.value)}
        />
    </div>
);

export default TextAreaField;
