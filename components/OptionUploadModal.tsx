
import React, { useState } from 'react';
import { Option, TravelRequest, RequestStatus, FlightDetails } from '../types';
import { gasService } from '../services/gasService';
import { ConfirmationDialog } from './ConfirmationDialog';

interface OptionUploadModalProps {
  request: TravelRequest;
  onClose: () => void;
  onSuccess: () => void;
}

const EMPTY_FLIGHT: FlightDetails = { airline: '', flightTime: '', notes: '' };

export const OptionUploadModal = ({ request, onClose, onSuccess }: OptionUploadModalProps) => {
  const [loading, setLoading] = useState(false);
  
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
  
  // Check if it's a round trip based on returnDate content
  const isRoundTrip = !!request.returnDate && request.returnDate.trim() !== '';

  const [options, setOptions] = useState<Option[]>([
    { 
      id: 'A', 
      totalPrice: 0, 
      outbound: { ...EMPTY_FLIGHT },
      inbound: isRoundTrip ? { ...EMPTY_FLIGHT } : undefined,
      hotel: request.requiresHotel ? '' : undefined,
      generalNotes: ''
    }
  ]);

  const handleFlightChange = (
    index: number, 
    leg: 'outbound' | 'inbound', 
    field: keyof FlightDetails, 
    value: string
  ) => {
    setOptions(prevOptions => prevOptions.map((opt, i) => {
        if (i !== index) return opt;

        // Ensure inbound object exists if needed
        let currentInbound = opt.inbound;
        if (leg === 'inbound' && !currentInbound) {
            currentInbound = { ...EMPTY_FLIGHT };
        }

        if (leg === 'outbound') {
            return { ...opt, outbound: { ...opt.outbound, [field]: value } };
        } else {
            return { ...opt, inbound: { ...currentInbound!, [field]: value } };
        }
    }));
  };

  const handleRootChange = (index: number, field: keyof Option, value: any) => {
    setOptions(prevOptions => prevOptions.map((opt, i) => {
        if (i !== index) return opt;
        return { ...opt, [field]: value };
    }));
  };

  const addOption = () => {
    const nextId = String.fromCharCode(65 + options.length); // A, B, C...
    setOptions([...options, { 
      id: nextId, 
      totalPrice: 0, 
      outbound: { ...EMPTY_FLIGHT },
      inbound: isRoundTrip ? { ...EMPTY_FLIGHT } : undefined,
      hotel: request.requiresHotel ? '' : undefined,
      generalNotes: ''
    }]);
  };

  const removeOption = (index: number) => {
    if (options.length > 1) {
      setOptions(options.filter((_, i) => i !== index));
    }
  };

  const validateOption = (opt: Option): { status: 'VALID' | 'EMPTY' | 'INCOMPLETE', missing?: string[] } => {
    // Check if fields have content
    const hasOutbound = !!opt.outbound.airline || !!opt.outbound.flightTime || !!opt.outbound.notes;
    const hasInbound = isRoundTrip ? (!!opt.inbound?.airline || !!opt.inbound?.flightTime || !!opt.inbound?.notes) : false;
    const hasHotel = request.requiresHotel ? !!opt.hotel : false;
    const hasPrice = opt.totalPrice > 0;
    
    // Consider an option "modified" if any string field is non-empty or price > 0
    const isModified = hasOutbound || hasInbound || hasHotel || hasPrice;

    if (!isModified) {
        return { status: 'EMPTY' };
    }

    // Identify Missing Required Fields
    const missing: string[] = [];

    // Outbound
    if (!opt.outbound.airline) missing.push('Aerolínea (Ida)');
    if (!opt.outbound.flightTime) missing.push('Hora Salida (Ida)');
    if (!opt.outbound.notes) missing.push('Notas/Escalas (Ida)');

    // Inbound (Only if Round Trip)
    if (isRoundTrip) {
        if (!opt.inbound?.airline) missing.push('Aerolínea (Regreso)');
        if (!opt.inbound?.flightTime) missing.push('Hora Salida (Regreso)');
        if (!opt.inbound?.notes) missing.push('Notas/Escalas (Regreso)');
    }

    // Hotel
    if (request.requiresHotel && !opt.hotel) {
        missing.push('Hotel');
    }

    if (missing.length === 0) return { status: 'VALID' };
    return { status: 'INCOMPLETE', missing };
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    
    const validOptions: Option[] = [];
    
    for (const opt of options) {
        const check = validateOption(opt);
        if (check.status === 'INCOMPLETE') {
            setDialog({
                isOpen: true,
                title: 'Opción Incompleta',
                message: `La Opción ${opt.id} está incompleta.\n\nFaltan los siguientes campos obligatorios:\n${check.missing?.join(', ')}\n\nPor favor complételos o elimine la opción.`,
                type: 'ALERT',
                onConfirm: closeDialog
            });
            return;
        }
        if (check.status === 'VALID') {
            validOptions.push(opt);
        }
    }

    if (validOptions.length === 0) {
        setDialog({
            isOpen: true,
            title: 'Sin Opciones',
            message: 'Debe registrar al menos una opción válida para continuar.',
            type: 'ALERT',
            onConfirm: closeDialog
        });
        return;
    }

    // Open Confirmation Dialog
    setDialog({
        isOpen: true,
        title: 'Confirmar Envío',
        message: `¿Está seguro de enviar ${validOptions.length} opción(es)?\n\nLa solicitud pasará a estado PENDIENTE_SELECCION y se notificará al usuario.`,
        type: 'CONFIRM',
        onConfirm: () => executeSubmission(validOptions),
        onCancel: closeDialog
    });
  };

  const executeSubmission = async (validOptions: Option[]) => {
      closeDialog(); // close confirmation
      setLoading(true);
      try {
        await gasService.updateRequestStatus(request.requestId, RequestStatus.PENDING_SELECTION, {
          analystOptions: validOptions
        });
        
        setDialog({
            isOpen: true,
            title: 'Envío Exitoso',
            message: 'Opciones enviadas correctamente.\n\nSe ha notificado al solicitante por correo electrónico para que realice su selección.',
            type: 'SUCCESS',
            onConfirm: () => {
                closeDialog();
                onSuccess();
            }
        });
      } catch (error) {
        setLoading(false);
        setDialog({
            isOpen: true,
            title: 'Error de Envío',
            message: 'Ocurrió un error al guardar las opciones: ' + error,
            type: 'ALERT',
            onConfirm: closeDialog
        });
      }
  };

  return (
    <>
      <ConfirmationDialog 
        isOpen={dialog.isOpen} 
        title={dialog.title} 
        message={dialog.message} 
        onConfirm={dialog.onConfirm} 
        onCancel={dialog.onCancel}
        type={dialog.type}
      />
      
      <div className="fixed inset-0 z-50 overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
        <div className="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
          
          <div className="fixed inset-0 bg-gray-500 bg-opacity-75 transition-opacity" onClick={onClose}></div>
          <span className="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>

          <div className="relative inline-block align-bottom bg-white rounded-lg px-4 pt-5 pb-4 text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-7xl sm:w-full sm:p-6">
            <div className="absolute top-0 right-0 pt-4 pr-4 z-10">
              <button onClick={onClose} className="bg-white rounded-md text-gray-400 hover:text-gray-500 text-2xl font-bold leading-none px-2 focus:outline-none">
                ✕
              </button>
            </div>

            <div className="w-full">
              <h3 className="text-lg leading-6 font-medium text-gray-900 mb-4 border-b pb-2">
                Cargar Opciones - Solicitud <span className="text-brand-red font-bold">{request.requestId}</span>
              </h3>
              
              <div className="flex flex-col lg:flex-row gap-6 h-[75vh]">
                
                {/* LEFT COLUMN: REQUEST DETAILS (READ ONLY) */}
                <div className="lg:w-1/3 bg-gray-50 rounded-lg p-4 overflow-y-auto border border-gray-200 shadow-inner">
                    <h4 className="text-xs font-bold text-gray-500 uppercase tracking-wider mb-4 border-b border-gray-200 pb-2">
                        Detalle Requerido
                    </h4>

                    {/* Route & Dates */}
                    <div className="mb-4">
                        <div className="grid grid-cols-2 gap-2 text-sm mb-2">
                            <div>
                                <span className="block text-xs text-gray-400">Origen</span>
                                <span className="font-bold text-gray-800">{request.origin}</span>
                            </div>
                            <div>
                                <span className="block text-xs text-gray-400">Destino</span>
                                <span className="font-bold text-gray-800">{request.destination}</span>
                            </div>
                        </div>
                        <div className="bg-white p-2 rounded border border-gray-200">
                             <div className="flex justify-between items-center mb-1">
                                <span className="text-xs text-gray-500">Ida:</span>
                                <span className="text-sm font-medium">{request.departureDate}</span>
                             </div>
                             {request.departureTimePreference && (
                                <div className="text-xs text-gray-400 text-right">Pref: {request.departureTimePreference}</div>
                             )}
                             
                             <div className="border-t my-1"></div>

                             <div className="flex justify-between items-center mb-1">
                                <span className="text-xs text-gray-500">Regreso:</span>
                                <span className="text-sm font-medium">{isRoundTrip ? request.returnDate : 'Solo Ida'}</span>
                             </div>
                             {isRoundTrip && request.returnTimePreference && (
                                <div className="text-xs text-gray-400 text-right">Pref: {request.returnTimePreference}</div>
                             )}
                        </div>
                    </div>

                    {/* Hotel Preference */}
                    {request.requiresHotel && (
                        <div className="mb-4 bg-blue-50 p-3 rounded border border-blue-100">
                             <span className="block text-xs font-bold text-blue-800 uppercase mb-1">Requiere Hotel ({request.nights} Noches)</span>
                             <div className="text-sm text-blue-900 font-medium">
                                {request.hotelName || 'Sin preferencia específica'}
                             </div>
                        </div>
                    )}

                    {/* User Notes (Highlighted) */}
                    {request.comments && (
                        <div className="mb-4 bg-yellow-50 p-3 rounded border border-yellow-200">
                            <span className="block text-xs font-bold text-yellow-800 uppercase mb-1">⚠️ Observaciones del Usuario</span>
                            <p className="text-sm text-gray-800 whitespace-pre-line italic">
                                "{request.comments}"
                            </p>
                        </div>
                    )}

                    {/* Passengers */}
                    <div className="mb-4">
                        <span className="block text-xs text-gray-400 uppercase mb-2">Pasajeros ({request.passengers.length})</span>
                        <div className="space-y-2">
                            {request.passengers.map((p, idx) => (
                                <div key={idx} className="text-xs bg-white p-2 rounded border border-gray-200 text-gray-700">
                                    <span className="font-bold block">{p.name}</span>
                                    <span className="text-gray-500">CC: {p.idNumber}</span>
                                </div>
                            ))}
                        </div>
                    </div>

                    {/* Cost Center info */}
                    <div className="mt-4 pt-4 border-t border-gray-200 text-xs text-gray-500">
                        <p><strong>Centro de Costos:</strong> {request.costCenter}</p>
                        <p><strong>Solicitante:</strong> {request.requesterEmail}</p>
                    </div>
                </div>

                {/* RIGHT COLUMN: FORM (EDITABLE) */}
                <div className="lg:w-2/3 flex flex-col h-full">
                  <div className="flex-1 overflow-y-auto pr-2">
                    <form id="optionsForm" onSubmit={handleSubmit} noValidate>
                        <div className="space-y-6">
                        {options.map((opt, idx) => (
                            <div key={idx} className="bg-white p-4 rounded-lg border-2 border-gray-200 relative shadow-sm">
                            <div className="absolute -top-3 -left-3 bg-brand-red text-white w-8 h-8 rounded-full flex items-center justify-center font-bold shadow">
                                {opt.id}
                            </div>
                            
                            {options.length > 1 && (
                                <button 
                                type="button"
                                onClick={() => removeOption(idx)}
                                className="absolute top-2 right-2 text-red-500 hover:text-red-700 text-xs font-bold border border-red-200 bg-red-50 px-2 py-1 rounded"
                                title="Eliminar Opción"
                                >
                                Eliminar
                                </button>
                            )}

                            <div className="mt-2 grid grid-cols-1 lg:grid-cols-2 gap-6">
                                {/* IDA */}
                                <div className="bg-gray-50 p-3 rounded border border-gray-100">
                                    <h5 className="text-xs font-bold text-gray-500 uppercase mb-3 border-b pb-1 flex items-center gap-2">
                                        <span className="text-lg">🛫</span> Vuelo de Ida
                                    </h5>
                                    <div className="space-y-3">
                                        <div>
                                            <label className="block text-xs font-medium text-gray-500">Aerolínea</label>
                                            <input
                                                type="text"
                                                className="mt-1 block w-full bg-white text-gray-900 placeholder-gray-500 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                placeholder="Ej: Avianca"
                                                value={opt.outbound.airline}
                                                onChange={(e) => handleFlightChange(idx, 'outbound', 'airline', e.target.value)}
                                            />
                                        </div>
                                        <div className="grid grid-cols-2 gap-2">
                                            <div>
                                                <label className="block text-xs font-medium text-gray-500">Hora Salida</label>
                                                <input
                                                    type="time"
                                                    className="mt-1 block w-full bg-white text-gray-900 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                    value={opt.outbound.flightTime}
                                                    onChange={(e) => handleFlightChange(idx, 'outbound', 'flightTime', e.target.value)}
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-medium text-gray-500">No. Vuelo</label>
                                                <input
                                                    type="text"
                                                    className="mt-1 block w-full bg-white text-gray-900 placeholder-gray-500 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                    placeholder="AV123"
                                                    value={opt.outbound.flightNumber || ''}
                                                    onChange={(e) => handleFlightChange(idx, 'outbound', 'flightNumber', e.target.value)}
                                                />
                                            </div>
                                        </div>
                                        <div>
                                            <label className="block text-xs font-medium text-gray-500">Notas / Escalas</label>
                                            <input
                                                type="text"
                                                className="mt-1 block w-full bg-white text-gray-900 placeholder-gray-500 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                placeholder="Ej: Directo, Maleta 23kg"
                                                value={opt.outbound.notes}
                                                onChange={(e) => handleFlightChange(idx, 'outbound', 'notes', e.target.value)}
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* VUELTA */}
                                {isRoundTrip && (
                                <div className="bg-gray-50 p-3 rounded border border-gray-100">
                                    <h5 className="text-xs font-bold text-gray-500 uppercase mb-3 border-b pb-1 flex items-center gap-2">
                                        <span className="text-lg">🛬</span> Vuelo de Regreso
                                    </h5>
                                    <div className="space-y-3">
                                        <div>
                                            <label className="block text-xs font-medium text-gray-500">Aerolínea</label>
                                            <input
                                                type="text"
                                                className="mt-1 block w-full bg-white text-gray-900 placeholder-gray-500 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                placeholder="Ej: Latam"
                                                value={opt.inbound?.airline || ''}
                                                onChange={(e) => handleFlightChange(idx, 'inbound', 'airline', e.target.value)}
                                            />
                                        </div>
                                        <div className="grid grid-cols-2 gap-2">
                                            <div>
                                                <label className="block text-xs font-medium text-gray-500">Hora Salida</label>
                                                <input
                                                    type="time"
                                                    className="mt-1 block w-full bg-white text-gray-900 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                    value={opt.inbound?.flightTime || ''}
                                                    onChange={(e) => handleFlightChange(idx, 'inbound', 'flightTime', e.target.value)}
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-xs font-medium text-gray-500">No. Vuelo</label>
                                                <input
                                                    type="text"
                                                    className="mt-1 block w-full bg-white text-gray-900 placeholder-gray-500 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                    placeholder="LA456"
                                                    value={opt.inbound?.flightNumber || ''}
                                                    onChange={(e) => handleFlightChange(idx, 'inbound', 'flightNumber', e.target.value)}
                                                />
                                            </div>
                                        </div>
                                        <div>
                                            <label className="block text-xs font-medium text-gray-500">Notas / Escalas</label>
                                            <input
                                                type="text"
                                                className="mt-1 block w-full bg-white text-gray-900 placeholder-gray-500 border-gray-300 rounded shadow-sm focus:ring-brand-red focus:border-brand-red text-sm p-1.5"
                                                placeholder="Ej: Escala en BOG 2h"
                                                value={opt.inbound?.notes || ''}
                                                onChange={(e) => handleFlightChange(idx, 'inbound', 'notes', e.target.value)}
                                            />
                                        </div>
                                    </div>
                                </div>
                                )}
                            </div>

                            {/* GENERAL (PRECIO Y HOTEL) */}
                            <div className="mt-4 pt-4 border-t border-gray-200 grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div>
                                    <label className="block text-xs font-bold text-gray-700 mb-1">Precio Total Opción (COP)</label>
                                    <div className="relative rounded-md shadow-sm">
                                        <div className="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3">
                                            <span className="text-gray-500 sm:text-sm">$</span>
                                        </div>
                                        <input
                                            type="number"
                                            className="block w-full rounded-md border-gray-300 pl-7 focus:border-brand-red focus:ring-brand-red sm:text-sm p-2 bg-yellow-50 font-bold text-gray-900"
                                            placeholder="0.00"
                                            value={opt.totalPrice}
                                            onChange={(e) => handleRootChange(idx, 'totalPrice', Number(e.target.value))}
                                        />
                                    </div>
                                </div>
                                
                                {request.requiresHotel && (
                                <div>
                                    <label className="block text-xs font-bold text-gray-700 mb-1">Hotel Incluido</label>
                                    <input
                                        type="text"
                                        className="block w-full bg-white text-gray-900 placeholder-gray-500 rounded-md border-gray-300 focus:border-brand-red focus:ring-brand-red sm:text-sm p-2"
                                        placeholder="Nombre del hotel"
                                        value={opt.hotel || ''}
                                        onChange={(e) => handleRootChange(idx, 'hotel', e.target.value)}
                                    />
                                </div>
                                )}
                            </div>
                            </div>
                        ))}
                        </div>
                        
                        <div className="mt-4 flex items-center justify-between">
                            <button
                                type="button"
                                onClick={addOption}
                                className="text-sm font-bold text-brand-red hover:text-red-800 flex items-center gap-1 border border-brand-red px-3 py-1 rounded-full hover:bg-red-50 transition"
                            >
                                <span>+ Añadir Opción</span>
                            </button>
                        </div>
                    </form>
                  </div>
                  
                  {/* Actions Footer - Fixed at bottom of right col */}
                  <div className="mt-4 pt-4 border-t bg-white flex justify-end gap-3">
                    <button
                        type="button"
                        onClick={onClose}
                        className="px-4 py-2 border border-gray-300 shadow-sm text-sm font-medium rounded-md text-gray-700 bg-white hover:bg-gray-50"
                        disabled={loading}
                    >
                        Cancelar
                    </button>
                    <button
                        // Trigger submit on the form programmatically or just move button inside form
                        onClick={(e) => handleSubmit(e as any)}
                        className="px-4 py-2 border border-transparent shadow-sm text-sm font-medium rounded-md text-white bg-brand-red hover:bg-red-700 focus:outline-none"
                        disabled={loading}
                    >
                        {loading ? 'Guardando...' : 'Enviar Opciones'}
                    </button>
                  </div>
                </div>

              </div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
};
