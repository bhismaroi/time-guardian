import type { CompiledEmployee } from '@/lib/types';
import { formatDateShort, isWeekend } from '@/lib/timeUtils';
import { cn } from '@/lib/utils';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';

interface AttendanceTableProps {
  employee: CompiledEmployee;
}

export function AttendanceTable({ employee }: AttendanceTableProps) {
  return (
    <div className="bg-card rounded-lg border overflow-hidden">
      <div className="px-4 py-3 border-b bg-muted/30">
        <div className="flex items-center justify-between">
          <div>
            <h3 className="font-semibold text-foreground">{employee.name}</h3>
            <p className="text-sm text-muted-foreground">NIP: {employee.nip}</p>
          </div>
          <div className="text-right text-sm text-muted-foreground">
            <p>{employee.division}</p>
            <p>{employee.department}</p>
          </div>
        </div>
      </div>
      <div className="overflow-x-auto">
        <Table>
          <TableHeader>
            <TableRow className="bg-muted/20">
              <TableHead className="w-20">Date</TableHead>
              <TableHead className="w-12">Day</TableHead>
              <TableHead className="w-20 text-center">Actual In</TableHead>
              <TableHead className="w-20 text-center">Actual Out</TableHead>
              <TableHead className="w-24 text-center">Total Hours</TableHead>
              <TableHead className="w-20 text-center">Tardiness</TableHead>
              <TableHead className="w-20 text-center">Leave Earlier</TableHead>
              <TableHead className="w-20 text-center">Overtime</TableHead>
              <TableHead className="w-28">Remarks</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {employee.records.map((record, index) => {
              const weekend = isWeekend(record.date);
              return (
                <TableRow 
                  key={index}
                  className={cn(
                    weekend && "bg-muted/30",
                    record.tardiness && "bg-destructive/5"
                  )}
                >
                  <TableCell className="font-medium">{formatDateShort(record.date)}</TableCell>
                  <TableCell className="text-muted-foreground">{record.dayOfWeek}</TableCell>
                  <TableCell className="text-center">{record.actualIn || '-'}</TableCell>
                  <TableCell className="text-center">{record.actualOut || '-'}</TableCell>
                  <TableCell className="text-center font-medium">
                    {weekend ? '-' : (record.totalHours || '0:00')}
                  </TableCell>
                  <TableCell className="text-center">
                    {record.tardiness && (
                      <span className="text-destructive font-medium">{record.tardiness}</span>
                    )}
                  </TableCell>
                  <TableCell className="text-center">
                    {record.leaveEarlier && (
                      <span className="text-warning font-medium">{record.leaveEarlier}</span>
                    )}
                  </TableCell>
                  <TableCell className="text-center">
                    {record.overtime && (
                      <span className="text-success font-medium">{record.overtime}</span>
                    )}
                  </TableCell>
                  <TableCell className="text-muted-foreground text-sm">{record.remarks}</TableCell>
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
