export interface DeliveryRecord {
  CLIENTE: string;
  CEDULA: string;
  NO_PROFORMA: string;
  CHASIS: string;
  FECHA_FACTURA: string;
  FECHA_ENTREGA: string;
  VENDEDOR: string;
  NO_TELEFONO: string;
  DIRECCION: string;
  DEPARTAMENTO: string;
  MODELO: string;
  COLOR: string;
  MOTOR: string;
  TIPO_VEHICULO: string;
  ANIO: string;
  KM_RECORRIDOS?: string;
  LLAVE_NO?: string;
  OBSERVACIONES?: string;
  BRAND?: 'isuzu' | 'gac';
  [key: string]: any; // Allow indexing
}

export type BrandType = 'isuzu' | 'gac';

export interface ColumnMapping {
  CLIENTE: string;
  CEDULA: string;
  NO_PROFORMA: string;
  CHASIS: string;
  FECHA_FACTURA: string;
  FECHA_ENTREGA: string;
  VENDEDOR: string;
  NO_TELEFONO: string;
  DIRECCION: string;
  DEPARTAMENTO: string;
  MODELO: string;
  COLOR: string;
  MOTOR: string;
  TIPO_VEHICULO: string;
  ANIO: string;
}

export interface ChecklistItem {
  id: string;
  label: string;
  checked: boolean;
}

export interface ChecklistSection {
  title: string;
  items: ChecklistItem[];
}
