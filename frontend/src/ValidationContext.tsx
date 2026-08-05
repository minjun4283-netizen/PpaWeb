import { createContext, useCallback, useContext, useState, type ReactNode } from "react";
import { api, type ValidationReport } from "./api";

interface ValidationContextValue {
  report: ValidationReport | null;
  running: boolean;
  run: () => Promise<void>;
}

const ValidationContext = createContext<ValidationContextValue | null>(null);

export function ValidationProvider({ children }: { children: ReactNode }) {
  const [report, setReport] = useState<ValidationReport | null>(null);
  const [running, setRunning] = useState(false);

  const run = useCallback(async () => {
    setRunning(true);
    try {
      const result = await api.runValidation();
      setReport(result);
    } finally {
      setRunning(false);
    }
  }, []);

  return (
    <ValidationContext.Provider value={{ report, running, run }}>
      {children}
    </ValidationContext.Provider>
  );
}

export function useValidation(): ValidationContextValue {
  const ctx = useContext(ValidationContext);
  if (!ctx) throw new Error("useValidation must be used inside ValidationProvider");
  return ctx;
}
