import * as React from "react";

interface AppLegacyProps {
  hostLabel: string;
}

export function AppLegacy({ hostLabel }: AppLegacyProps): React.ReactElement {
  return (
    <div
      style={{
        padding: 16,
        fontFamily: '"Segoe UI", Tahoma, sans-serif',
        color: "#242424",
        lineHeight: 1.45,
      }}
    >
      <h2 style={{ margin: "0 0 8px", fontSize: 18 }}>Word AI Chat</h2>
      <p style={{ margin: "0 0 8px" }}>
        <strong>IE11 host OK</strong> — Webpack ES5 bundle loaded successfully.
      </p>
      <p style={{ margin: "0 0 8px" }}>Host: {hostLabel}</p>
      <p style={{ margin: 0, color: "#616161", fontSize: 13 }}>
        Phase IE-0 proof of life. Full UI port follows in IE-1 through IE-4.
      </p>
    </div>
  );
}