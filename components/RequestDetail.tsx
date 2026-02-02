
import React, { useState } from 'react';
import { TravelRequest, Option, FlightDetails, RequestStatus, Integrant } from '../types';
import { gasService } from '../services/gasService';
import { ConfirmationDialog } from './ConfirmationDialog';
import { ModificationForm } from './ModificationForm';

interface RequestDetailProps {
  request: TravelRequest;
  integrantes: Integrant[]; // Added
  onClose: () => void;
  onRefresh?: () => void;
}

const FlightLegDetail = ({ leg, title }: { leg?: FlightDetails, title: string }) => {
  if (!leg || !leg.airline) return null;
  return (
    <div className="mb-2 text-sm border-l-2 border-gray-300 pl-2">
      <strong className="text-gray-600 block text-xs uppercase">{title}</strong>
      <div className="flex justify-between">
        <span>{leg.airline} {leg.flightNumber ? `(${leg.flightNumber})` : ''}</span>
        <span className="font-mono">{leg.flightTime}</span>
      </div>
      <p className="text-xs text-gray-500 italic">{leg.notes}</p>
    </div>
  );
};

interface OptionCardProps {
  option: Option;
  isSelected?: boolean;
}

const OptionCard: React.FC<OptionCardProps> = ({ option, isSelected = false }) => (
  <div className={`p-3 rounded-md border ${isSelected ? 'border-green-500 bg-green-50 ring-1 ring-green-500' : 'border-gray-200 bg-white'}`}>
    <div className="flex justify-between items-center mb-2 border-b pb-1">
      <span className={`font-bold ${isSelected ? 'text-green-700' : 'text-gray-700'}`}>
        Opción {option.id} {isSelected && '(SELECCIONADA)'}
      </span>
      <span className="font-bold text-gray-900">$ {Number(option.totalPrice).toLocaleString()}</span>
    </div>
    <FlightLegDetail leg={option.outbound} title="Ida" />
    <FlightLegDetail leg={option.inbound} title="Regreso" />
    {option.hotel && (
      <div className="mt-2 text-xs text-blue-800 bg-blue-50 p-1 rounded">
        <strong>Hotel:</strong> {option.hotel}
      </div>
    )}
  </div>
);

export const RequestDetail = ({ request, integrantes, onClose, onRefresh }: RequestDetailProps) => {
  const [loading, setLoading] = useState(false);
  const [showModifyForm, setShowModifyForm] = useState(false);
  
  // Dialog State
  const [dialog, setDialog] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    type: 'ALERT' | 'CONFIRM' | 'SUCCESS';
    onConfirm: () => void;
    onCancel?: () => void;
  }>({
    isOpen: false,
    title: '',
    message: '',
    type: 'ALERT',
    onConfirm: () => {},
  });

  const closeDialog = () => setDialog({ ...dialog, isOpen: false });

  const hasOptions = !!(request.analystOptions && request.analystOptions.length > 0);
  const hasSelection = !!request.selectedOption;
  const isOneWay = !request.returnDate;

  // Logic for Finalize Button
  const canFinalize = request.status === RequestStatus.APPROVED && request.supportData && request.supportData.files.length > 0;
  
  // Logic for Modification
  // Can modify if NOT processed (closed) and NOT already pending a change
  const canModify = request.status !== RequestStatus.PROCESSED && request.status !== RequestStatus.PENDING_CHANGE_APPROVAL;

  const handleFinalizeClick = () => {
    setDialog({
        isOpen: true,
        title: 'Finalizar Solicitud',
        message: '¿Está seguro de cerrar esta solicitud definitivamente?',
        type: 'CONFIRM',
        onConfirm: executeFinalize,
        onCancel: closeDialog
    });
  };

  const executeFinalize = async () => {
    closeDialog();
    setLoading(true);
    try {
        await gasService.closeRequest(request.requestId);
        
        setDialog({
            isOpen: true,
            title: 'Solicitud Cerrada',
            message: 'El proceso se ha completado exitosamente.',
            type: 'SUCCESS',
            onConfirm: () => {
                closeDialog();
                if (onRefresh) onRefresh();
                onClose(); // Close modal
            }
        });
    } catch (e) {
        setDialog({
            isOpen: true,
            title: 'Error',
            message: "Error: " + e,
            type: 'ALERT',
            onConfirm: closeDialog
        });
    } finally {
        setLoading(false);
    }
  };

  return (
    <>
      <ConfirmationDialog 
        isOpen={dialog.isOpen}
        title={dialog.title}
        message={dialog.message}
        type={dialog.type}
        onConfirm={dialog.onConfirm}
        onCancel={dialog.onCancel}
      />

      {showModifyForm && (
        <ModificationForm 
            originalRequest={request} 
            integrantes={integrantes} // Added
            onClose={() => setShowModifyForm(false)}
            onSuccess={() => {
                setShowModifyForm(false);
                setDialog({
                    isOpen: true,
                    title: 'Petición Enviada',
                    message: 'Su solicitud de modificación ha sido enviada al administrador para revisión.',
                    type: 'SUCCESS',
                    onConfirm: () => {
                        closeDialog();
                        if (onRefresh) onRefresh();
                        onClose();
                    }
                });
            }}
        />
      )}

      <div className="fixed inset-0 z-[60] overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
        <div className="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
          
          <div 
            className="fixed inset-0 bg-gray-500 bg-opacity-75 transition-opacity" 
            aria-hidden="true" 
            onClick={onClose}
          ></div>

          <span className="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>

          <div className="relative inline-block align-bottom bg-white rounded-lg px-4 pt-5 pb-4 text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-2xl sm:w-full sm:p-6">
            <div className="absolute top-0 right-0 pt-4 pr-4 z-10">
              <button
                type="button"
                className="bg-white rounded-md text-gray-400 hover:text-gray-500 text-2xl font-bold leading-none px-2 focus:outline-none"
                onClick={onClose}
              >
                <span className="sr-only">Cerrar</span>
                ✕
              </button>
            </div>
            
            <div className="w-full">
              <div className="border-b border-gray-200 pb-4 mb-4">
                  <div className="flex flex-wrap items-center gap-3">
                      <h3 className="text-xl leading-6 font-bold text-gray-900">
                          Detalle Solicitud <span className="text-brand-red">{request.requestId}</span>
                      </h3>
                      <div className={`px-3 py-1 text-xs font-bold rounded-full ${
                          request.status === 'APROBADO' ? 'bg-green-100 text-green-700' : 
                          request.status === 'DENEGADO' ? 'bg-red-100 text-red-700' : 
                          request.status === 'PROCESADO' ? 'bg-gray-800 text-white' :
                          request.status === 'PENDIENTE_APROBACION_CAMBIO' ? 'bg-purple-100 text-purple-700' :
                          'bg-gray-100 text-gray-600'
                      }`}>
                          {request.status.replace(/_/g, ' ')}
                      </div>
                      
                      {request.hasChangeFlag && (
                          <span className="px-3 py-1 text-xs font-bold rounded-full bg-yellow-100 text-yellow-800 animate-pulse">
                              CAMBIO GENERADO
                          </span>
                      )}
                  </div>
                  <p className="text-xs text-gray-500 mt-1">Solicitado el: {new Date(request.timestamp).toLocaleString()}</p>
                  
                  {request.status === RequestStatus.PENDING_CHANGE_APPROVAL && (
                       <div className="mt-3 bg-purple-50 p-2 rounded border border-purple-200 text-xs text-purple-800">
                           <strong>🔄 Modificación en Proceso:</strong> Esta solicitud tiene cambios pendientes de aprobación por parte del administrador.
                           <br/><em className="text-purple-600 mt-1 block">"{request.changeReason}"</em>
                       </div>
                  )}

                  {/* Mostramos el motivo del cambio si ya fue aprobado (CAMBIO GENERADO) */}
                  {request.hasChangeFlag && request.changeReason && request.status !== RequestStatus.PENDING_CHANGE_APPROVAL && (
                       <div className="mt-3 bg-yellow-50 p-2 rounded border border-yellow-200 text-xs text-yellow-800">
                           <strong>📝 Motivo del Último Cambio:</strong> {request.changeReason}
                       </div>
                  )}
              </div>
              
              <div className="space-y-6 max-h-[70vh] overflow-y-auto pr-2">
                  
                  {/* 1. INFORMACIÓN DE LA SOLICITUD ORIGINAL */}
                  <section>
                      <div className="flex justify-between items-center mb-3">
                          <h4 className="text-xs font-bold text-gray-400 uppercase tracking-wider">1. Requerimiento Inicial</h4>
                          
                          {/* BOTÓN SOLICITAR CAMBIO */}
                          {canModify && (
                              <button 
                                onClick={() => setShowModifyForm(true)}
                                className="text-xs bg-gray-100 hover:bg-gray-200 text-gray-700 px-3 py-1 rounded border border-gray-300 font-bold transition flex items-center gap-1"
                              >
                                  ✏️ Solicitar Cambio
                              </button>
                          )}
                      </div>
                      
                      <div className="bg-gray-50 rounded-lg p-4 grid grid-cols-1 md:grid-cols-2 gap-4 text-sm relative">
                          {request.hasChangeFlag && (
                              <div className="absolute top-0 right-0 p-1">
                                  <span className="text-[10px] bg-yellow-200 text-yellow-800 px-1 rounded font-bold">INFO ACTUALIZADA</span>
                              </div>
                          )}

                          {/* Ruta y Fechas */}
                          <div className="col-span-2 grid grid-cols-2 gap-4 border-b border-gray-200 pb-3 mb-1">
                              <div>
                                  <span className="block text-xs text-gray-500">Origen</span>
                                  <span className="font-semibold text-gray-900">{request.origin}</span>
                              </div>
                              <div>
                                  <span className="block text-xs text-gray-500">Destino</span>
                                  <span className="font-semibold text-gray-900">{request.destination}</span>
                              </div>
                              <div>
                                  <span className="block text-xs text-gray-500">Ida</span>
                                  <span className="font-medium">{request.departureDate}</span>
                                  {request.departureTimePreference && <span className="text-xs text-gray-500 ml-1">({request.departureTimePreference})</span>}
                              </div>
                              <div>
                                  <span className="block text-xs text-gray-500">Regreso</span>
                                  {isOneWay ? (
                                      <span className="text-xs italic text-gray-400">Solo Ida</span>
                                  ) : (
                                      <>
                                          <span className="font-medium">{request.returnDate}</span>
                                          {request.returnTimePreference && <span className="text-xs text-gray-500 ml-1">({request.returnTimePreference})</span>}
                                      </>
                                  )}
                              </div>
                          </div>

                          {/* Pasajeros */}
                          <div className="col-span-2">
                              <span className="block text-xs text-gray-500 mb-1">Pasajeros ({request.passengers.length})</span>
                              <div className="flex flex-wrap gap-2">
                                  {request.passengers.map((p, i) => (
                                      <span key={i} className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-white border border-gray-300 text-gray-700">
                                          {p.name} ({p.idNumber})
                                      </span>
                                  ))}
                              </div>
                          </div>

                          {/* Datos Admin */}
                          <div>
                              <span className="block text-xs text-gray-500">Empresa / Sede</span>
                              <span className="font-medium block truncate">
                                  {request.company} / {request.site}
                              </span>
                          </div>
                          <div>
                              <span className="block text-xs text-gray-500">Centro de Costos</span>
                              <span className="font-medium block truncate" title={request.costCenter}>
                                  {request.costCenter === 'VARIOS' ? 'VARIOS' : request.costCenter}
                              </span>
                              {request.variousCostCenters && (
                                  <p className="text-xs text-gray-400 truncate">{request.variousCostCenters}</p>
                              )}
                          </div>
                          <div>
                              <span className="block text-xs text-gray-500">Aprobador Asignado</span>
                              <span className="font-medium text-brand-red text-xs block truncate" title={request.approverEmail}>
                                  {request.approverEmail || 'No asignado'}
                              </span>
                          </div>
                          
                          {/* Hotel */}
                          {request.requiresHotel && (
                              <div className="col-span-2 bg-blue-50 p-2 rounded border border-blue-100 mt-1">
                                  <div className="flex justify-between">
                                      <span className="text-xs font-bold text-blue-700 uppercase">Hospedaje:</span>
                                      <span className="text-xs font-bold text-blue-700">({request.nights} Noches)</span>
                                  </div>
                                  <span className="text-sm text-blue-900 ml-2">{request.hotelName}</span>
                              </div>
                          )}

                          {/* Observaciones */}
                          {request.comments && (
                              <div className="col-span-2 bg-yellow-50 p-3 rounded border border-yellow-200 mt-1">
                                  <span className="block text-xs font-bold text-yellow-800 mb-1">Observaciones / Notas:</span>
                                  <p className="text-sm text-gray-700 whitespace-pre-line">{request.comments}</p>
                              </div>
                          )}
                      </div>
                  </section>

                  {/* 2. OPCIÓN SELECCIONADA (SI EXISTE) */}
                  {hasSelection && request.selectedOption && (
                      <section>
                          <h4 className="text-xs font-bold text-green-600 uppercase tracking-wider mb-3">2. Opción Seleccionada por Usuario</h4>
                          <OptionCard option={request.selectedOption} isSelected={true} />
                      </section>
                  )}

                  {/* 3. OPCIONES DISPONIBLES (SI AÚN NO SELECCIONA O HISTÓRICO) */}
                  {hasOptions && !hasSelection && (
                      <section>
                          <h4 className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-3">2. Opciones Propuestas</h4>
                          <div className="space-y-3">
                              {request.analystOptions?.map((opt, idx) => (
                                  <OptionCard key={idx} option={opt} />
                              ))}
                          </div>
                      </section>
                  )}

                  {!hasOptions && !hasSelection && (
                      <div className="text-center py-6 bg-gray-50 border-2 border-dashed border-gray-200 rounded-lg">
                          <p className="text-gray-400 text-sm">Aún no se han cargado opciones de cotización.</p>
                      </div>
                  )}
              </div>
            </div>
            
            <div className="mt-5 sm:mt-6 border-t pt-4 sm:grid sm:grid-cols-2 sm:gap-3 sm:grid-flow-row-dense">
              {canFinalize ? (
                  <>
                      <button
                          type="button"
                          className="w-full inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 bg-green-600 text-base font-medium text-white hover:bg-green-700 focus:outline-none disabled:opacity-50 sm:col-start-2 sm:text-sm"
                          onClick={handleFinalizeClick}
                          disabled={loading}
                      >
                          {loading ? 'Procesando...' : 'Finalizar Solicitud'}
                      </button>
                      <button
                          type="button"
                          className="mt-3 w-full inline-flex justify-center rounded-md border border-gray-300 shadow-sm px-4 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 focus:outline-none sm:mt-0 sm:col-start-1 sm:text-sm"
                          onClick={onClose}
                      >
                          Cerrar
                      </button>
                  </>
              ) : (
                  <button
                      type="button"
                      className="col-span-2 w-full inline-flex justify-center rounded-md border border-gray-300 shadow-sm px-4 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 focus:outline-none sm:text-sm"
                      onClick={onClose}
                  >
                      Cerrar
                  </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </>
  );
};
