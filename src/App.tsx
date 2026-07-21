/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Mail, CheckCircle, FileSpreadsheet, AlertCircle, LogOut, Paperclip, FileText } from 'lucide-react';
import { initAuth, googleSignIn, logout, getAccessToken } from './lib/google-auth';
import { initMicrosoftAuth, microsoftSignIn, getMicrosoftAccessToken, microsoftLogout } from './lib/microsoft-auth';
import type { User } from 'firebase/auth';
import { sendEmail } from './lib/gmail-service';
import { sendOutlookEmail } from './lib/outlook-service';

interface RowData {
  trato: string;
  nombre: string;
  cargo: string;
  institucion: string;
  correo: string;
  adjunto: string;
  estado?: 'pendiente' | 'enviado' | 'error';
}

export default function App() {
  const [data, setData] = useState<RowData[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [needsAuth, setNeedsAuth] = useState(true);
  const [user, setUser] = useState<any>(null); // Use any for unified user representation
  const [authProvider, setAuthProvider] = useState<'google' | 'microsoft' | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [attachment, setAttachment] = useState<File | null>(null);
  const [subject, setSubject] = useState('');
  const [explicacion, setExplicacion] = useState('');
  const [showConfirm, setShowConfirm] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');

  useEffect(() => {
    // Try to initialize both auth providers on mount
    const tryInit = async () => {
      let isMicrosoftLoggedIn = false;
      // First try Microsoft (since it uses MSAL session storage, it might be active)
      await initMicrosoftAuth(
        (user, token) => {
          setUser(user);
          setAuthProvider('microsoft');
          setNeedsAuth(false);
          isMicrosoftLoggedIn = true;
        }
      );
      
      // If Microsoft didn't log us in, try Google
      if (!isMicrosoftLoggedIn) {
        initAuth(
          (user, token) => {
            setUser(user);
            setAuthProvider('google');
            setNeedsAuth(false);
          },
          () => {
            // Neither logged in
          }
        );
      }
    };
    
    tryInit();
  }, [needsAuth]); // Re-run if we need auth

  const handleGoogleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setUser(result.user);
        setAuthProvider('google');
        setNeedsAuth(false);
      }
    } catch (err) {
      console.error('Google Login failed:', err);
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleMicrosoftLogin = async () => {
    setIsLoggingIn(true);
    try {
      await microsoftSignIn();
      // Code won't reach here normally as loginRedirect redirects the page.
    } catch (err) {
      console.error('Microsoft Login failed:', err);
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    if (authProvider === 'google') {
      await logout();
    } else if (authProvider === 'microsoft') {
      await microsoftLogout();
    }
    setAuthProvider(null);
    setUser(null);
    setNeedsAuth(true);
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert sheet to JSON, treating the first row as headers
        const rows = XLSX.utils.sheet_to_json<any>(worksheet, { header: 1 });
        
        // Skip header row and map
        const mappedData: RowData[] = rows.slice(1).map((row) => ({
          trato: String(row[0] || ''),
          nombre: String(row[1] || ''),
          cargo: String(row[2] || ''),
          institucion: String(row[3] || ''),
          correo: String(row[4] || ''),
          adjunto: String(row[5] || ''),
          estado: 'pendiente'
        }));
        setData(mappedData);
      };
      reader.readAsBinaryString(file);
    } catch (error) {
      console.error("Error reading file:", error);
      setErrorMessage("Hubo un error al leer el archivo Excel.");
    }
  };

  const generarCuerpo = (row: RowData) => {
    const estimado = row.trato.toLowerCase().trim() === 'señora' ? 'Estimada' : (row.trato.toLowerCase().trim() === 'señor' ? 'Estimado' : 'Estimada/o');
    return `${row.trato}
${row.nombre}
${row.cargo}
${row.institucion}

${estimado} ${row.trato}:

${explicacion}`;
  };

  const handleEnvio = async () => {
    if (!showConfirm) {
      setShowConfirm(true);
      return;
    }

    setShowConfirm(false);
    setIsProcessing(true);
    setErrorMessage('');
    
    let token = null;
    if (authProvider === 'google') {
      token = await getAccessToken();
    } else if (authProvider === 'microsoft') {
      token = await getMicrosoftAccessToken();
    }

    if (!token) {
      setErrorMessage("No hay sesión activa. Por favor inicie sesión nuevamente.");
      setNeedsAuth(true);
      setIsProcessing(false);
      return;
    }

    // Send emails one by one
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.estado === 'enviado') continue;

      try {
        const body = generarCuerpo(row);
        const nombreAdjunto = attachment ? attachment.name : row.adjunto;
        const finalSubject = subject || `Oficio adjunto: ${nombreAdjunto}`;
        
        if (authProvider === 'google') {
          await sendEmail(token, row.correo, finalSubject, body, attachment);
        } else if (authProvider === 'microsoft') {
          await sendOutlookEmail(token, row.correo, finalSubject, body, attachment);
        }
        
        setData(prev => {
          const newData = [...prev];
          newData[i].estado = 'enviado';
          return newData;
        });
      } catch (error) {
        console.error(`Error sending email to ${row.correo}:`, error);
        setData(prev => {
          const newData = [...prev];
          newData[i].estado = 'error';
          return newData;
        });
      }
    }
    
    setIsProcessing(false);
  };

  if (needsAuth) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 max-w-md w-full text-center space-y-6">
          <div className="w-16 h-16 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mx-auto">
            <Mail size={32} />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-gray-900">Iniciar Sesión</h1>
            <p className="mt-2 text-gray-500">
              Para enviar correos, necesita autorizar la aplicación con su cuenta de Google o Microsoft Outlook.
            </p>
          </div>
          
          <div className="space-y-3">
            <button 
              onClick={handleGoogleLogin}
              disabled={isLoggingIn}
              className="w-full relative flex items-center justify-center bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 font-medium py-2.5 px-4 rounded-lg transition-colors shadow-sm disabled:opacity-50"
            >
              <svg className="w-5 h-5 mr-3" viewBox="0 0 48 48">
                <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                <path fill="none" d="M0 0h48v48H0z"></path>
              </svg>
              {isLoggingIn ? 'Iniciando sesión...' : 'Continuar con Google'}
            </button>

            <button 
              onClick={handleMicrosoftLogin}
              disabled={isLoggingIn}
              className="w-full relative flex items-center justify-center bg-[#0078D4] hover:bg-[#106EBE] text-white font-medium py-2.5 px-4 rounded-lg transition-colors shadow-sm disabled:opacity-50"
            >
              <svg className="w-5 h-5 mr-3" viewBox="0 0 23 23" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="1" y="1" width="10" height="10" fill="#F25022"/>
                <rect x="12" y="1" width="10" height="10" fill="#7FBA00"/>
                <rect x="1" y="12" width="10" height="10" fill="#00A4EF"/>
                <rect x="12" y="12" width="10" height="10" fill="#FFB900"/>
              </svg>
              {isLoggingIn ? 'Iniciando sesión...' : 'Continuar con Outlook'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-7xl mx-auto space-y-8">
        
        <div className="text-center">
          <h1 className="text-3xl font-bold text-gray-900 tracking-tight">
            Automatización de Correos
          </h1>
          <p className="mt-2 text-lg text-gray-600">
            Cargue su base de datos y revise los correos a enviar.
          </p>
          <div className="mt-4 flex flex-col sm:flex-row items-center justify-center gap-4">
            <div className="flex items-center gap-2 text-sm text-gray-600 bg-white px-4 py-2 rounded-full border border-gray-200 shadow-sm">
              <span className="w-2 h-2 rounded-full bg-green-500"></span>
              Conectado como {user?.email}
            </div>
            <button 
              onClick={handleLogout}
              className="text-sm font-medium text-gray-500 hover:text-gray-700 flex items-center gap-1 transition-colors"
            >
              <LogOut size={16} /> Cerrar Sesión
            </button>
          </div>
        </div>

        {/* Upload Card */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 flex flex-col items-center justify-center space-y-4">
            <div className="w-16 h-16 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center">
              <FileSpreadsheet size={32} />
            </div>
            <div className="text-center">
              <h3 className="text-lg font-medium text-gray-900">1. Subir Base de Datos Excel</h3>
              <p className="mt-1 text-sm text-gray-500 max-w-sm">
                Columnas requeridas: Trato, Nombre, Cargo, Institución, Correo_Destino y Nombre_Adjunto.
              </p>
            </div>
            <label className="relative cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-6 py-3 rounded-lg font-medium transition-colors shadow-sm">
              <span>Seleccionar Archivo .xlsx</span>
              <input 
                type="file" 
                className="sr-only" 
                accept=".xlsx"
                onChange={handleFileUpload}
              />
            </label>
            {data.length > 0 && (
              <span className="text-sm font-medium text-green-600 flex items-center gap-1">
                <CheckCircle size={16} /> {data.length} filas cargadas
              </span>
            )}
          </div>

          <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 flex flex-col items-center justify-center space-y-4">
            <div className="w-16 h-16 bg-purple-50 text-purple-600 rounded-full flex items-center justify-center">
              <Paperclip size={32} />
            </div>
            <div className="text-center">
              <h3 className="text-lg font-medium text-gray-900">2. Archivo Adjunto (Opcional)</h3>
              <p className="mt-1 text-sm text-gray-500 max-w-sm">
                Sube un archivo que se adjuntará a todos los correos electrónicos.
              </p>
            </div>
            <label className="relative cursor-pointer bg-purple-600 hover:bg-purple-700 text-white px-6 py-3 rounded-lg font-medium transition-colors shadow-sm">
              <span>Seleccionar Archivo Común</span>
              <input 
                type="file" 
                className="sr-only" 
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) {
                    setAttachment(file);
                    setSubject(`Oficio adjunto: ${file.name}`);
                  }
                }}
              />
            </label>
            {attachment && (
              <span className="text-sm font-medium text-green-600 flex items-center gap-1">
                <CheckCircle size={16} /> {attachment.name}
              </span>
            )}
          </div>
        </div>

        {/* Asunto (Común para todos) */}
        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100">
          <h3 className="text-lg font-medium text-gray-900 flex items-center gap-2 mb-4">
            <Mail className="text-gray-400" size={20} />
            3. Asunto del Correo
          </h3>
          <input
            type="text"
            value={subject}
            onChange={(e) => setSubject(e.target.value)}
            placeholder="Asunto del correo (se autocompleta al subir el archivo adjunto)..."
            className="w-full p-4 text-sm border border-gray-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
          />
        </div>

        {/* Explicación (Común para todos) */}
        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100">
          <h3 className="text-lg font-medium text-gray-900 flex items-center gap-2 mb-4">
            <FileText className="text-gray-400" size={20} />
            4. Desarrollo o Explicación (Común para todos)
          </h3>
          <textarea
            value={explicacion}
            onChange={(e) => setExplicacion(e.target.value)}
            placeholder="Escriba aquí el desarrollo, contexto o explicación del archivo que se incluirá en todos los correos..."
            className="w-full h-32 p-4 text-sm border border-gray-200 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none resize-none"
          ></textarea>
        </div>

        {/* Data Table */}
        {data.length > 0 && (
          <div className="space-y-6">
            <div className="bg-white shadow-sm border border-gray-100 rounded-2xl overflow-hidden">
              <div className="px-6 py-5 border-b border-gray-100 flex items-center justify-between">
                <h3 className="text-lg font-medium text-gray-900 flex items-center gap-2">
                  <Mail className="text-gray-400" size={20} />
                  Vista Previa de Correos ({data.length})
                </h3>
              </div>
              <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Estado</th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Destinatario</th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Institución</th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Correo</th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Mensaje Generado</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {data.map((row, idx) => (
                    <tr key={idx} className="hover:bg-gray-50 transition-colors">
                      <td className="px-6 py-4 whitespace-nowrap">
                        {row.estado === 'pendiente' ? (
                          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-yellow-100 text-yellow-800">
                            Pendiente
                          </span>
                        ) : row.estado === 'error' ? (
                          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-800 gap-1">
                            <AlertCircle size={12} /> Error
                          </span>
                        ) : (
                          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800 gap-1">
                            <CheckCircle size={12} /> Enviado
                          </span>
                        )}
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap">
                        <div className="text-sm font-medium text-gray-900">{row.nombre}</div>
                        <div className="text-sm text-gray-500">{row.cargo}</div>
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                        {row.institucion}
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                        {row.correo}
                      </td>
                      <td className="px-6 py-4">
                        <div className="text-sm text-gray-500 whitespace-pre-wrap min-w-80 bg-gray-50 p-3 rounded-lg border border-gray-100 font-mono text-xs">
                          {generarCuerpo(row)}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Action Button */}
          <div className="flex flex-col items-end pt-4 pb-12 gap-4">
            {errorMessage && (
              <div className="text-red-600 bg-red-50 px-4 py-2 rounded-lg border border-red-200 flex items-center gap-2">
                <AlertCircle size={18} />
                <span>{errorMessage}</span>
              </div>
            )}
            <button
              onClick={handleEnvio}
              disabled={isProcessing}
              className={`${showConfirm ? 'bg-red-600 hover:bg-red-700' : 'bg-blue-600 hover:bg-blue-700'} disabled:opacity-50 disabled:cursor-not-allowed text-white px-8 py-4 rounded-xl font-bold text-lg transition-colors shadow-lg flex items-center gap-3`}
            >
              <Mail size={24} />
              {isProcessing ? 'Enviando Correos...' : showConfirm ? `Confirmar envío de ${data.length} correos` : '5. Enviar Correos'}
            </button>
            {showConfirm && (
              <button
                onClick={() => setShowConfirm(false)}
                className="text-gray-500 hover:text-gray-700 text-sm font-medium"
              >
                Cancelar
              </button>
            )}
          </div>
        </div>
        )}

      </div>
    </div>
  );
}

