import * as React from "react";

export interface IeSelectOption {
  value: string;
  label: string;
}

interface IeSelectProps {
  label?: string;
  value: string;
  options: IeSelectOption[];
  disabled?: boolean;
  className?: string;
  fieldClassName?: string;
  onChange: (value: string) => void;
}

export function IeSelect(props: IeSelectProps): React.ReactElement {
  var className = "ie-select";
  if (props.className) {
    className += " " + props.className;
  }

  var fieldClassName = "ie-select-field";
  if (props.fieldClassName) {
    fieldClassName += " " + props.fieldClassName;
  }

  return (
    <div className={fieldClassName}>
      {props.label ? <label className="ie-select-label">{props.label}</label> : null}
      <select
        className={className}
        value={props.value}
        disabled={props.disabled}
        onChange={function (event) {
          props.onChange(event.target.value);
        }}
      >
        {props.options.map(function (option) {
          return (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          );
        })}
      </select>
    </div>
  );
}