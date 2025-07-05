export interface Ticket {
  name: string;
  phone: string;
  email: string;
  seat: {
    row: string;
    number: string;
  };
  uniqueCode: string;
  checkedInTime: Date | null;
}
