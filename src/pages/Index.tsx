import { FileUpload } from '@/components/FileUpload';
import { EmployeeTabs } from '@/components/EmployeeTabs';
import { useAttendanceCompiler } from '@/hooks/useAttendanceCompiler';
import { Button } from '@/components/ui/button';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Clock3, Download, RotateCcw, Loader2, AlertCircle } from 'lucide-react';

const Index = () => {
  const {
    fingerprintFile,
    setFingerprintFile,
    onlineFile,
    setOnlineFile,
    compiledData,
    isCompiling,
    error,
    canCompile,
    hasData,
    compile,
    downloadReport,
    reset,
  } = useAttendanceCompiler();

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b bg-card">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-primary flex items-center justify-center">
              <Clock3 className="w-5 h-5 text-primary-foreground" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-foreground">Time Guardian</h1>
              <p className="text-sm text-muted-foreground">Compile attendance from fingerprint and online sources</p>
            </div>
          </div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-8">
        <div className="grid md:grid-cols-2 gap-4 mb-6">
          <FileUpload
            label="Fingerprint Excel"
            description="Upload the fingerprint attendance file"
            file={fingerprintFile}
            onFileSelect={setFingerprintFile}
          />
          <FileUpload
            label="Online Excel"
            description="Upload the online attendance file"
            file={onlineFile}
            onFileSelect={setOnlineFile}
          />
        </div>

        {error && (
          <Alert variant="destructive" className="mb-6">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}

        <div className="flex flex-wrap gap-3 mb-8">
          <Button onClick={compile} disabled={!canCompile || isCompiling} size="lg" className="min-w-32">
            {isCompiling ? (
              <>
                <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                Compiling...
              </>
            ) : (
              <>
                <ClipboardCheck className="w-4 h-4 mr-2" />
                Compile
              </>
            )}
          </Button>
          <Button onClick={downloadReport} disabled={!hasData} variant="secondary" size="lg" className="min-w-32">
            <Download className="w-4 h-4 mr-2" />
            Download Report
          </Button>
          <Button onClick={reset} variant="outline" size="lg" className="min-w-32">
            <RotateCcw className="w-4 h-4 mr-2" />
            Reset
          </Button>
        </div>

        {hasData && (
          <section>
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-foreground">
                Compiled Attendance ({compiledData.length} employees)
              </h2>
            </div>
            <EmployeeTabs employees={compiledData} />
          </section>
        )}

        {!hasData && !isCompiling && (
          <div className="text-center py-16">
            <div className="w-16 h-16 rounded-full bg-muted mx-auto flex items-center justify-center mb-4">
              <Clock3 className="w-8 h-8 text-muted-foreground" />
            </div>
            <h3 className="text-lg font-medium text-foreground mb-2">No Data Yet</h3>
            <p className="text-muted-foreground max-w-md mx-auto">
              Upload both the Fingerprint and Online Excel files, then click Compile to generate the attendance report.
            </p>
          </div>
        )}

        <section className="mt-12 p-6 bg-card rounded-lg border">
          <h3 className="font-semibold text-foreground mb-4">Calculation Rules</h3>
          <div className="grid md:grid-cols-2 gap-6 text-sm text-muted-foreground">
            <div>
              <h4 className="font-medium text-foreground mb-2">Break Deductions</h4>
              <ul className="space-y-1">
                <li>- Monday to Thursday: 12:00 - 12:30 (30 min)</li>
                <li>- Friday: 11:30 - 13:00 (90 min)</li>
              </ul>
            </div>
            <div>
              <h4 className="font-medium text-foreground mb-2">Flexi Time</h4>
              <ul className="space-y-1">
                <li>- Flexi 1: 08:00 - 08:15, out by 16:45 (Fri: 17:15)</li>
                <li>- Flexi 2: 08:15 - 08:30, out by 17:00 (Fri: 17:30)</li>
                <li>- After 08:30 = Tardiness</li>
              </ul>
            </div>
            <div>
              <h4 className="font-medium text-foreground mb-2">Overtime</h4>
              <ul className="space-y-1">
                <li>- Monday to Thursday: starts from 17:30</li>
                <li>- Friday: starts from 18:00</li>
              </ul>
            </div>
            <div>
              <h4 className="font-medium text-foreground mb-2">Data Merging</h4>
              <ul className="space-y-1">
                <li>- Earliest clock-in from both sources</li>
                <li>- Latest clock-out from both sources</li>
              </ul>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
};

export default Index;
