import React from "react";

interface Props {
    label: string;
    value: string;
    placeholder?: string;
    type?: string;
    onChange: (value: string) => void;
}

const InputField: React.FC<Props> = ({
    label,
    value,
    placeholder,
    type = "text",
    onChange,
}) => (
    <div className="mb-3">
        <label className="form-label">{label}</label>
        <input
            type={type}
            className="form-control"
            placeholder={placeholder}
            value={value}
            onChange={(e) => onChange(e.target.value)}
        />
    </div>
);

export default InputField;
