import { ChecklistSection } from './types';

export const FIELD_ALIASES_MAP = {
  CLIENTE:       ['nombre_cliente','cliente','nombre','name','customer','razon_social'],
  CEDULA:        ['cedula','no_cedula','nro_cedula','cedula_ruc','ruc','identificacion','id_cliente'],
  NO_PROFORMA:   ['no_factura','factura','invoice','proforma','num_factura','numero_factura','nro_factura','no_proforma','numero_proforma'],
  CHASIS:        ['chasis','chasis_prs_fi','vin','chassis','num_chasis','numero_chasis','no_chasis','serie'],
  FECHA_FACTURA: ['fecha_documento','fecha_factura','fecha','date','fecha_doc','fecha_emision'],
  FECHA_ENTREGA: ['fecha_de_entrega','fecha_entrega','delivery_date','entrega','fecha_despacho'],
  VENDEDOR:      ['vendedor','seller','asesor','agente','sales','asesor_ventas','cod_vendedor'],
  NO_TELEFONO:   ['no_telefono','telefono','phone','celular','tel','movil'],
  DIRECCION:     ['direccion','address','dir','ubicacion','domicilio'],
  DEPARTAMENTO:  ['departamento','department','ciudad','region','municipio'],
  MODELO:        ['modelo','model','descripcion','articulo','codigo_articulo','desc_articulo','vehiculo'],
  COLOR:         ['color','colour','colorvehiculo'],
  MOTOR:         ['motor','engine','no_motor','num_motor','numero_motor'],
  TIPO_VEHICULO: ['tipo','tipo_vehiculo','vehicle_type','categoria','tipo_unidad'],
  ANIO:          ['anio','año','year','modelo_anio','anio_modelo']
};

export const DEFAULT_CHECKLIST: ChecklistSection[] = [
  {
    title: 'REVISIONES EXTERIORES',
    items: [
      { id: 'ext1', label: 'Revision de carroceria', checked: true },
      { id: 'ext2', label: 'Espejos', checked: true },
      { id: 'ext3', label: 'Tapon de Combustible', checked: true },
      { id: 'ext4', label: 'Antena', checked: true },
      { id: 'ext5', label: 'Puertas', checked: true },
      { id: 'ext6', label: 'Rines y copas', checked: true },
    ]
  },
  {
    title: 'REVISIONES INTERIORES',
    items: [
      { id: 'int1', label: 'Indicadores de tableros', checked: true },
      { id: 'int2', label: 'Indicadores de advertencia', checked: true },
      { id: 'int3', label: 'Encendido de focos / techo / mapa', checked: true },
      { id: 'int4', label: 'Luces de parqueo', checked: true },
      { id: 'int5', label: 'Parabrisa / enjuague', checked: true },
      { id: 'int6', label: 'Marcha de retroceso', checked: true },
      { id: 'int7', label: 'Porta vasos', checked: true },
      { id: 'int8', label: 'Tapizado / Cinturones', checked: true },
      { id: 'int9', label: 'Logos y Marcas', checked: true },
      { id: 'int10', label: 'Tapa de Fusible / Guantera', checked: true },
      { id: 'int11', label: 'Aire acondicionado', checked: true },
    ]
  },
  {
    title: 'REVISION DE LUCES',
    items: [
      { id: 'luc1', label: 'Focos frontales', checked: true },
      { id: 'luc2', label: 'Lucen en general', checked: true },
      { id: 'luc3', label: 'Retroceso', checked: true },
      { id: 'luc4', label: 'Emergencia', checked: true },
    ]
  },
  {
    title: 'REVISION DE MOTOR',
    items: [
      { id: 'mot1', label: 'Medidor de aceite', checked: true },
      { id: 'mot2', label: 'Liquido limpia parablrisas', checked: true },
      { id: 'mot3', label: 'Liquido Hidraulico', checked: true },
      { id: 'mot4', label: 'QA', checked: true },
      { id: 'mot5', label: 'Bateria', checked: true },
      { id: 'mot6', label: 'Fusibles', checked: true },
    ]
  },
  {
    title: 'REVISION DE VALIJERO',
    items: [
      { id: 'val1', label: 'Apertura de balijero', checked: true },
      { id: 'val2', label: 'Apertura de tapon de combustible', checked: true },
      { id: 'val3', label: 'Alfombra', checked: true },
      { id: 'val4', label: 'Llanta de repuesto', checked: true },
      { id: 'val5', label: 'Gata / Maneral / Llave de rueda', checked: true },
      { id: 'val6', label: 'Luz Interior', checked: true },
    ]
  },
  {
    title: 'ENTREGA DE DOCUMENTOS',
    items: [
      { id: 'doc1', label: 'Manual de conductor', checked: true },
      { id: 'doc2', label: 'Factura comercial', checked: true },
      { id: 'doc3', label: 'Garantia de Fabrica', checked: true },
      { id: 'doc4', label: 'Constancia para fines de transito', checked: true },
      { id: 'doc5', label: 'Libreta de servicio', checked: true },
      { id: 'doc6', label: 'Emision de Gases', checked: true },
    ]
  }
];

export const ACCESSORIES_LIST = [
  { id: 'acc_gata', label: 'GATA Y MANERAL' },
  { id: 'acc_repuesto', label: 'LLANTA DE REPUESTO' },
  { id: 'acc_manual', label: 'MANUAL DEL CONDUCTOR' },
  { id: 'acc_ac', label: 'AIRE ACONDICIONADO' },
  { id: 'acc_espejos_lat', label: 'ESPEJOS LATERALES' },
  { id: 'acc_herramientas', label: 'HERRAMIENTAS' },
  { id: 'acc_radio', label: 'RADIO' },
  { id: 'acc_espejo_int', label: 'ESPEJO INTERIOR' },
  { id: 'acc_km', label: 'KM RECORRIDOS' },
];

export const GENERAL_OBSERVATIONS_TEMPLATES = [
  "Este vehículo deberá ser enviado al TALLER DE SERVICIO, que está situado en nuestras instalaciones en Managua, para efectuar dos chequeos de garantía LIBRE DE COSTO DE MANO DE OBRA, a los 1500 Km. y 5000 Km. de recorrido.\nGARANTIA DE ISUZU ES DE 100,000KMS O TRES AÑOS DE USO, LO QUE SE CUMPLA PRIMERO.\nNOTA: -Filtros y lubricantes son por cuenta del cliente.\nRECIBO CONDICIONES DE GARANTÍA___________________",
  "Unidad recibida con protectores de asientos originales. Se entrega juego doble de llaves operativas. Próxima revisión recomendada en Taller SAMSA a los 1,500 km.",
  "Se incluye manual del operador original y póliza de garantía física. Vehículo entregado con rines pulidos de fábrica.",
  "Garantía extendida aplicable de 3 años o 100,000 kilómetros. Cliente capacitado brevemente sobre el panel de testigos virtuales GAC Intelligent Drive."
];
