
import React, { useState, useEffect } from 'react';
import { TravelRequest, Passenger, RequestStatus, CostCenterMaster, Integrant } from '../types';
import { COMPANIES, SITES, MAX_PASSENGERS } from '../constants';
import { gasService } from '../services/gasService';

interface RequestFormProps {
  userEmail: string;
  integrantes: Integrant[]; // Data passed from App
  onSuccess: () => void;
  onCancel: () => void;
}

type TripType = 'ROUND_TRIP' | 'ONE_WAY';

export const RequestForm: React.FC<RequestFormProps> = ({ userEmail, integrantes, onSuccess, onCancel }) => {
  const [loading, setLoading] = useState(false);
  const [passengers, setPassengers] = useState<Passenger[]>([{ name: '', idNumber: '', email: '' }]);
  
  // Trip Logic
  const [tripType, setTripType] = useState<TripType>('ROUND_TRIP');
  
  // Hotel Logic
  const [requiresHotel, setRequiresHotel] = useState(false);
  const [manualNights, setManualNights] = useState<boolean>(false); // Checkbox for manual override
  const [numberOfNights, setNumberOfNights] = useState<number>(0);

  // Master Data State
  const [masterData, setMasterData] = useState<CostCenterMaster[]>([]);
  const [availableBusinessUnits, setAvailableBusinessUnits] = useState<string[]>([]);
  const [filteredCostCenters, setFilteredCostCenters] = useState<CostCenterMaster[]>([]);
  
  // Various Cost Center State
  const [variousCCList, setVariousCCList] = useState<string[]>([]);
  const [variousCCInput, setVariousCCInput] = useState<string>('');

  // Basic State
  const [formData, setFormData] = useState<Partial<TravelRequest>>({
    company: '',
    businessUnit: '',
    site: '',
    costCenter: '',
    origin: '',
    destination: '',
    departureDate: '',
    returnDate: '',
    departureTimePreference: '',
    returnTimePreference: '',
    workOrder: '',
    hotelName: '',
    comments: '', // New Notes Field
  });

  // Load Master Data on Mount
  useEffect(() => {
    const fetchMasters = async () => {
      try {
        const data = await gasService.getCostCenterData();
        setMasterData(data);
        
        // Extract Unique Business Units
        const uniqueUnits = Array.from(new Set(data.map(item => item.businessUnit)))
                             .filter(u => u && u !== 'NA') 
                             .sort();
        setAvailableBusinessUnits(uniqueUnits);
      } catch (err) {
        console.error("Error loading masters:", err);
      }
    };
    fetchMasters();
  }, []);

  // Filter Cost Centers when Business Unit Changes
  useEffect(() => {
    if (formData.businessUnit) {
      const filtered = masterData.filter(item => item.businessUnit === formData.businessUnit);
      const variosOption: CostCenterMaster = { code: 'VARIOS', name: 'Múltiples Centros de Costo', businessUnit: formData.businessUnit };
      setFilteredCostCenters([...filtered, variosOption]);
    } else {
      setFilteredCostCenters([]);
    }
  }, [formData.businessUnit, masterData]);

  // Handle Trip Type Change
  useEffect(() => {
    if (tripType === 'ONE_WAY') {
        setManualNights(true);
        setFormData(prev => ({ ...prev, returnDate: '', returnTimePreference: '' }));
    } else {
        setManualNights(false);
    }
  }, [tripType]);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    let finalValue = value;

    if (name === 'origin' || name === 'destination') {
       finalValue = value
        .toUpperCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") 
        .replace(/[^A-Z\s]/g, ""); 
    }
    
    if (name === 'businessUnit') {
      setFormData(prev => ({ ...prev, [name]: finalValue, costCenter: '' }));
      setVariousCCList([]); 
    } else {
      setFormData(prev => ({ ...prev, [name]: finalValue }));
    }
  };

  const handleOpenPicker = (e: React.MouseEvent<HTMLInputElement>) => {
    try {
      if ('showPicker' in e.currentTarget) {
        e.currentTarget.showPicker();
      }
    } catch (error) {}
  };

  const handlePassengerChange = (index: number, field: keyof Passenger, value: string) => {
    let finalValue = value;

    if (field === 'name') {
      finalValue = value
        .toUpperCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/[^A-Z\s]/g, "");
    } else if (field === 'idNumber') {
      finalValue = value.replace(/[^0-9]/g, "");
    }

    const newPassengers = [...passengers];
    newPassengers[index] = { ...newPassengers[index], [field]: finalValue };

    // --- AUTO-FILL LOGIC ---
    if (field === 'idNumber') {
        const found = integrantes.find(i => i.idNumber === finalValue);
        if (found) {
            newPassengers[index].name = found.name;
            newPassengers[index].email = found.email;
        } else {
            // If user deletes ID or changes to one not found, clear/enable name?
            // Usually safer to keep current text if typing, but if pasting ID we want accurate.
            // Let's clear name only if it was previously auto-filled (hard to track)
            // Strategy: Just update email to empty if not found, but keep name manual
            if (newPassengers[index].email && !newPassengers[index].email.includes('@')) {
               // was likely empty or partial
            }
        }
    }
    // -----------------------

    setPassengers(newPassengers);
  };

  const addPassenger = () => {
    if (passengers.length < MAX_PASSENGERS) {
      setPassengers([...passengers, { name: '', idNumber: '', email: '' }]);
    }
  };

  const removePassenger = (index: number) => {
    if (passengers.length > 1) {
      setPassengers(passengers.filter((_, i) => i !== index));
    }
  };

  const handleAddVariousCC = () => {
    if (!variousCCInput.trim()) return;
    let code = variousCCInput.trim();
    if (/^\d+$/.test(code)) {
      code = code.padStart(4, '0');
    }
    if (!variousCCList.includes(code)) {
      setVariousCCList([...variousCCList, code]);
    }
    setVariousCCInput('');
  };

  const handleRemoveVariousCC = (codeToRemove: string) => {
    setVariousCCList(variousCCList.filter(c => c !== codeToRemove));
  };

  // Calculate Nights Automatically for Round Trip
  useEffect(() => {
    if (requiresHotel && tripType === 'ROUND_TRIP' && !manualNights && formData.departureDate && formData.returnDate) {
        const d1 = new Date(formData.departureDate);
        const d2 = new Date(formData.returnDate);
        const diffTime = Math.abs(d2.getTime() - d1.getTime());
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        setNumberOfNights(diffDays > 0 ? diffDays : 0);
    }
  }, [requiresHotel, tripType, manualNights, formData.departureDate, formData.returnDate]);

  // Check if passenger exists in database to disable/enable Name field
  const isPassengerInDb = (idNumber: string) => {
      return integrantes.some(i => i.idNumber === idNumber);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (formData.costCenter === 'VARIOS' && variousCCList.length === 0) {
      alert('Debe agregar al menos un centro de costos en la lista de VARIOS.');
      return;
    }

    if (tripType === 'ROUND_TRIP' && (!formData.returnDate)) {
        alert('Para vuelos de ida y regreso, la fecha de retorno es obligatoria.');
        return;
    }

    if (requiresHotel && numberOfNights <= 0) {
        alert('El número de noches de hospedaje debe ser mayor a 0.');
        return;
    }

    setLoading(true);

    try {
      const payload: Partial<TravelRequest> = {
        ...formData,
        returnDate: tripType === 'ONE_WAY' ? '' : formData.returnDate,
        returnTimePreference: tripType === 'ONE_WAY' ? '' : formData.returnTimePreference,
        requesterEmail: userEmail,
        passengers,
        requiresHotel,
        nights: requiresHotel ? numberOfNights : 0,
        timestamp: new Date().toISOString(),
        status: RequestStatus.PENDING_OPTIONS,
        variousCostCenters: formData.costCenter === 'VARIOS' ? variousCCList.join(', ') : undefined
      };

      await gasService.createRequest(payload);
      onSuccess();
    } catch (error) {
      alert('Error creando solicitud: ' + error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="bg-white shadow-lg rounded-lg overflow-hidden">
      <div className="bg-gray-50 px-6 py-4 border-b border-gray-200">
        <h2 className="text-lg font-medium text-gray-900">Nueva Solicitud de Viaje</h2>
        <p className="mt-1 text-sm text-gray-500">Diligencie todos los campos obligatorios.</p>
      </div>

      <form onSubmit={handleSubmit} className="p-6 space-y-8">
        
        {/* Section 1: General Info */}
        <div className="space-y-6">
            
            {/* Row 1: Company & Site */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700">Empresa *</label>
                <select 
                  name="company" 
                  required
                  className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500"
                  value={formData.company}
                  onChange={handleInputChange}
                >
                  <option value="">Seleccione...</option>
                  {COMPANIES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Sede *</label>
                <select 
                  name="site" 
                  required
                  className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500"
                  value={formData.site}
                  onChange={handleInputChange}
                >
                  <option value="">Seleccione...</option>
                  {SITES.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            </div>

            {/* Row 2: Business/WorkOrder vs CostCenter */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-start">
              
              <div className="space-y-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700">Unidad de Negocio *</label>
                  <select 
                    name="businessUnit" 
                    required
                    className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500"
                    value={formData.businessUnit}
                    onChange={handleInputChange}
                    disabled={availableBusinessUnits.length === 0}
                  >
                    <option value="">
                      {availableBusinessUnits.length === 0 ? 'Cargando...' : 'Seleccione...'}
                    </option>
                    {availableBusinessUnits.map(u => <option key={u} value={u}>{u}</option>)}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700">Orden de Trabajo (Opcional)</label>
                  <input
                    type="text"
                    name="workOrder"
                    className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500"
                    value={formData.workOrder}
                    onChange={handleInputChange}
                  />
                </div>
              </div>

              <div className="space-y-6">
                <div>
                  <label className="block text-sm font-medium text-gray-700">Centro de Costos *</label>
                  <select 
                    name="costCenter" 
                    required
                    className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500"
                    value={formData.costCenter}
                    onChange={handleInputChange}
                    disabled={!formData.businessUnit || filteredCostCenters.length === 0}
                  >
                    <option value="">
                      {!formData.businessUnit 
                        ? 'Seleccione Unidad Primero' 
                        : filteredCostCenters.length === 0 
                          ? 'No hay centros asociados' 
                          : 'Seleccione...'}
                    </option>
                    {filteredCostCenters.map(c => (
                      <option key={c.code} value={c.code}>
                        {c.code === 'VARIOS' ? 'VARIOS - Múltiples Centros' : `${c.code} - ${c.name || 'Sin descripción'}`}
                      </option>
                    ))}
                  </select>

                  {formData.costCenter === 'VARIOS' && (
                    <div className="mt-3 bg-gray-50 p-3 rounded-md border border-gray-200">
                      <label className="block text-xs font-bold text-gray-700 mb-2">Ingrese los centros de costos:</label>
                      <div className="flex gap-2 mb-2">
                        <input
                          type="text"
                          className="flex-1 rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-1 bg-white text-gray-900"
                          placeholder="Ej: 0101"
                          value={variousCCInput}
                          onChange={(e) => setVariousCCInput(e.target.value)}
                          onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), handleAddVariousCC())}
                        />
                        <button
                          type="button"
                          onClick={handleAddVariousCC}
                          className="bg-brand-red text-white text-xs px-3 py-1 rounded font-bold hover:bg-red-700"
                        >
                          Agregar
                        </button>
                      </div>
                      
                      {variousCCList.length > 0 && (
                        <div className="flex flex-wrap gap-2 mt-2">
                          {variousCCList.map((cc, idx) => (
                            <span key={idx} className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-gray-100 text-gray-800 border border-gray-300">
                              {cc}
                              <button
                                type="button"
                                onClick={() => handleRemoveVariousCC(cc)}
                                className="ml-1.5 inline-flex flex-shrink-0 h-4 w-4 rounded-full text-gray-400 hover:bg-gray-200 hover:text-gray-500 focus:outline-none"
                              >
                                <span className="sr-only">Remove</span>
                                <svg className="h-2 w-2" stroke="currentColor" fill="none" viewBox="0 0 8 8">
                                  <path strokeLinecap="round" strokeWidth="1.5" d="M1 1l6 6m0-6L1 7" />
                                </svg>
                              </button>
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>

            </div>
        </div>

        <hr />

        {/* Section 2: Passengers */}
        <div>
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-md font-medium text-gray-900">Pasajeros ({passengers.length})</h3>
            {passengers.length < MAX_PASSENGERS && (
              <button 
                type="button" 
                onClick={addPassenger}
                className="text-sm text-brand-red font-semibold hover:text-red-700"
              >
                + Agregar Pasajero
              </button>
            )}
          </div>
          <div className="space-y-4">
            {passengers.map((p, idx) => (
              <div key={idx} className="flex flex-col sm:flex-row gap-4 items-start sm:items-end bg-gray-50 p-4 rounded-md">
                <div className="flex-1 w-full sm:w-auto">
                  <label className="block text-xs font-medium text-gray-500">Cédula (P{idx + 1}) *</label>
                  <input
                    type="text"
                    required
                    className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500"
                    value={p.idNumber}
                    onChange={(e) => handlePassengerChange(idx, 'idNumber', e.target.value)}
                    placeholder="Escriba Cédula"
                  />
                </div>
                <div className="flex-1 w-full sm:w-auto">
                  <label className="block text-xs font-medium text-gray-500">Nombre Completo *</label>
                  <input
                    type="text"
                    required
                    readOnly={isPassengerInDb(p.idNumber)} // Disable if found
                    className={`mt-1 block w-full rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500 ${isPassengerInDb(p.idNumber) ? 'bg-gray-200 cursor-not-allowed text-gray-600' : 'bg-white'}`}
                    value={p.name}
                    onChange={(e) => handlePassengerChange(idx, 'name', e.target.value)}
                    placeholder="Nombre"
                  />
                </div>
                {/* Hidden email input field logic if needed, but we just store it in state */}
                {passengers.length > 1 && (
                  <button 
                    type="button" 
                    onClick={() => removePassenger(idx)}
                    className="text-red-500 p-2 hover:bg-red-50 rounded"
                  >
                    🗑️
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>

        <hr />

        {/* Section 3: Itinerary */}
        <div>
           <div className="flex justify-between items-center mb-4">
             <h3 className="text-md font-medium text-gray-900">Itinerario</h3>
             
             {/* TRIP TYPE TOGGLE */}
             <div className="flex bg-gray-200 p-1 rounded-lg">
                <button
                    type="button"
                    onClick={() => setTripType('ROUND_TRIP')}
                    className={`px-3 py-1 text-xs font-bold rounded-md transition ${tripType === 'ROUND_TRIP' ? 'bg-white shadow text-brand-red' : 'text-gray-500 hover:text-gray-700'}`}
                >
                    Ida y Regreso
                </button>
                <button
                    type="button"
                    onClick={() => setTripType('ONE_WAY')}
                    className={`px-3 py-1 text-xs font-bold rounded-md transition ${tripType === 'ONE_WAY' ? 'bg-white shadow text-brand-red' : 'text-gray-500 hover:text-gray-700'}`}
                >
                    Solo Ida
                </button>
             </div>
           </div>

           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-gray-700">Ciudad Origen *</label>
                <input
                  type="text"
                  name="origin"
                  required
                  className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 uppercase text-gray-900 placeholder-gray-500"
                  value={formData.origin}
                  onChange={handleInputChange}
                  placeholder="EJ: BOGOTA"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Ciudad Destino *</label>
                <input
                  type="text"
                  name="destination"
                  required
                  className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 uppercase text-gray-900 placeholder-gray-500"
                  value={formData.destination}
                  onChange={handleInputChange}
                  placeholder="EJ: CALI"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Fecha Ida *</label>
                <input
                  type="date"
                  name="departureDate"
                  required
                  style={{ colorScheme: 'light' }}
                  className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500 cursor-pointer"
                  value={formData.departureDate}
                  onChange={handleInputChange}
                  onClick={handleOpenPicker}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Hora Ida (Pref.)</label>
                <input
                  type="time"
                  name="departureTimePreference"
                  style={{ colorScheme: 'light' }}
                  className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500 cursor-pointer"
                  value={formData.departureTimePreference}
                  onChange={handleInputChange}
                  onClick={handleOpenPicker}
                />
              </div>
              
              {/* Return Fields - Conditional */}
              {tripType === 'ROUND_TRIP' && (
                  <>
                    <div>
                        <label className="block text-sm font-medium text-gray-700">Fecha Vuelta *</label>
                        <input
                        type="date"
                        name="returnDate"
                        required={tripType === 'ROUND_TRIP'}
                        style={{ colorScheme: 'light' }}
                        className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500 cursor-pointer"
                        value={formData.returnDate}
                        onChange={handleInputChange}
                        onClick={handleOpenPicker}
                        />
                    </div>
                    <div>
                        <label className="block text-sm font-medium text-gray-700">Hora Vuelta (Pref.)</label>
                        <input
                        type="time"
                        name="returnTimePreference"
                        style={{ colorScheme: 'light' }}
                        className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-500 cursor-pointer"
                        value={formData.returnTimePreference}
                        onChange={handleInputChange}
                        onClick={handleOpenPicker}
                        />
                    </div>
                  </>
              )}
           </div>
        </div>

        <hr />

        {/* Section 4: Hotel */}
        <div>
          <div className="flex items-center gap-3 mb-4">
             <div className="flex items-center h-5">
                <input
                  id="hotel"
                  type="checkbox"
                  className="focus:ring-brand-red h-4 w-4 text-brand-red border-gray-300 rounded"
                  checked={requiresHotel}
                  onChange={(e) => setRequiresHotel(e.target.checked)}
                />
             </div>
             <div className="text-sm">
                <label htmlFor="hotel" className="font-medium text-gray-700">¿Requiere Hospedaje?</label>
             </div>
          </div>

          {requiresHotel && (
            <div className="bg-blue-50 p-4 rounded-md border border-blue-100 space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700">Nombre del Hotel (Preferencia) *</label>
                <input
                    type="text"
                    name="hotelName"
                    required={requiresHotel}
                    className="mt-1 block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 uppercase text-gray-900 placeholder-gray-500"
                    value={formData.hotelName}
                    onChange={handleInputChange}
                    placeholder="Escriba su hotel de preferencia..."
                />
              </div>
              
              {/* Hotel Nights Logic */}
              <div className="bg-white p-3 rounded border border-gray-200">
                
                {tripType === 'ROUND_TRIP' && (
                    <div className="flex items-center gap-2 mb-2">
                        <input
                            id="manualNights"
                            type="checkbox"
                            className="focus:ring-brand-red h-4 w-4 text-brand-red border-gray-300 rounded"
                            checked={manualNights}
                            onChange={(e) => setManualNights(e.target.checked)}
                        />
                        <label htmlFor="manualNights" className="text-xs text-gray-700 font-bold">
                            ¿Las fechas de hospedaje son diferentes a las del vuelo?
                        </label>
                    </div>
                )}

                {manualNights || tripType === 'ONE_WAY' ? (
                    <div>
                        <label className="block text-sm font-medium text-gray-700">Número de Noches *</label>
                        <input
                            type="number"
                            min="1"
                            required
                            className="mt-1 block w-32 bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900"
                            value={numberOfNights}
                            onChange={(e) => setNumberOfNights(parseInt(e.target.value) || 0)}
                        />
                        {tripType === 'ONE_WAY' && (
                             <p className="text-xs text-blue-600 mt-1">Para vuelos "Solo Ida", debe indicar manualmente las noches.</p>
                        )}
                    </div>
                ) : (
                    <div>
                        <span className="text-sm text-gray-600">Noches calculadas según fechas de vuelo: </span>
                        <span className="font-bold text-gray-900 text-lg">{numberOfNights}</span>
                    </div>
                )}
              </div>
            </div>
          )}
        </div>

        <hr />

        {/* Section 5: Observations / Notes */}
        <div>
           <label className="block text-sm font-medium text-gray-700 mb-1">Observaciones / Notas Adicionales</label>
           <textarea
             name="comments"
             rows={3}
             className="block w-full bg-white rounded-md border-gray-300 shadow-sm focus:border-brand-red focus:ring-brand-red sm:text-sm border p-2 text-gray-900 placeholder-gray-400"
             placeholder="Escriba aquí requerimientos de equipaje, detalles sobre el horario o preferencias específicas de vuelo..."
             value={formData.comments}
             onChange={handleInputChange}
           />
        </div>

        {/* Actions */}
        <div className="flex justify-end gap-3 pt-4">
          <button
            type="button"
            onClick={onCancel}
            className="px-4 py-2 border border-gray-300 shadow-sm text-sm font-medium rounded-md text-gray-700 bg-white hover:bg-gray-50 focus:outline-none"
            disabled={loading}
          >
            Cancelar
          </button>
          <button
            type="submit"
            className="px-4 py-2 border border-transparent shadow-sm text-sm font-medium rounded-md text-white bg-brand-red hover:bg-red-700 focus:outline-none disabled:opacity-50"
            disabled={loading}
          >
            {loading ? 'Procesando...' : 'Crear Solicitud'}
          </button>
        </div>

      </form>
    </div>
  );
};
