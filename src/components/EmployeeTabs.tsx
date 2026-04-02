import { useState } from 'react';
import type { CompiledEmployee } from '@/lib/types';
import { AttendanceTable } from './AttendanceTable';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { ScrollArea, ScrollBar } from '@/components/ui/scroll-area';

interface EmployeeTabsProps {
  employees: CompiledEmployee[];
}

export function EmployeeTabs({ employees }: EmployeeTabsProps) {
  const [activeTab, setActiveTab] = useState(employees[0]?.sheetName || '');

  if (employees.length === 0) {
    return null;
  }

  return (
    <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
      <ScrollArea className="w-full whitespace-nowrap">
        <TabsList className="inline-flex h-10 items-center justify-start rounded-md bg-muted p-1 mb-4">
          {employees.map((employee) => (
            <TabsTrigger
              key={employee.sheetName}
              value={employee.sheetName}
              className="px-3 py-1.5 text-sm font-medium whitespace-nowrap"
            >
              {employee.name}
            </TabsTrigger>
          ))}
        </TabsList>
        <ScrollBar orientation="horizontal" />
      </ScrollArea>
      {employees.map((employee) => (
        <TabsContent key={employee.sheetName} value={employee.sheetName} className="mt-0">
          <AttendanceTable employee={employee} />
        </TabsContent>
      ))}
    </Tabs>
  );
}
