export interface Report {
  id: string;
  requesterName: string;
  campus: string;
  importDate: string;
  exportDate: string;
  items: Record<string, number>;
  status: 'Process' | 'Done';
}