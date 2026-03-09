import React, { useState, useEffect, useMemo, useCallback } from 'react';
import msruasLogo from './assets/msruas_logo.png';
import logo from './assets/logo1.png';
// ==========================================
// 1. MICROSOFT 365 CONFIGURATION
// ==========================================
const MSAL_CONFIG = {
  auth: {
    clientId: "c4e24997-2d1d-468d-a975-548ca8019c56", 
    // "common" allows anyone with a Microsoft institutional account to log in
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin, 
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  }
};

// Files.ReadWrite.All allows the logged-in user to write to your shared folder
const LOGIN_REQUEST = {
  scopes: ["User.Read", "Files.ReadWrite"] 
};

let msalInstance = null;

// Add your HOD's actual emails here
const HOD_EMAILS = [
  "23etcs002030@msruas.ac.in",
  "chiraggk8@gmail.com"
];

const determineUserRole = (account) => {
  if (!account || !account.username) return 'FACULTY';
  const email = account.username.toLowerCase();
  return HOD_EMAILS.map(e => e.toLowerCase()).includes(email) ? 'HOD' : 'FACULTY';
};

// ==========================================
// 2. ICONS, BRANDING & UI ASSETS
// ==========================================

const IconFileText = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>;
const IconSettings = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>;
const IconBookOpen = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"></path><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"></path></svg>;
const IconTarget = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="6"></circle><circle cx="12" cy="12" r="2"></circle></svg>;
const IconCheckCircle = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>;
const IconPlus = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>;
const IconTrash2 = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>;
const IconList = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="8" y1="6" x2="21" y2="6"></line><line x1="8" y1="12" x2="21" y2="12"></line><line x1="8" y1="18" x2="21" y2="18"></line><line x1="3" y1="6" x2="3.01" y2="6"></line><line x1="3" y1="12" x2="3.01" y2="12"></line><line x1="3" y1="18" x2="3.01" y2="18"></line></svg>;
const IconActivity = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>;
const IconGrid = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="7" height="7"></rect><rect x="14" y="3" width="7" height="7"></rect><rect x="14" y="14" width="7" height="7"></rect><rect x="3" y="14" width="7" height="7"></rect></svg>;
const IconDownload = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>;
const IconLogOut = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path><polyline points="16 17 21 12 16 7"></polyline><line x1="21" y1="12" x2="9" y2="12"></line></svg>;
const IconEye = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>;
const IconEdit = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>;
const IconArrowLeft = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="19" y1="12" x2="5" y2="12"></line><polyline points="12 19 5 12 12 5"></polyline></svg>;
const IconCloud = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"></path></svg>;
const IconMicrosoft = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="8" height="8" fill="#F25022"/><rect x="13" y="3" width="8" height="8" fill="#7FBA00"/><rect x="3" y="13" width="8" height="8" fill="#00A4EF"/><rect x="13" y="13" width="8" height="8" fill="#FFB900"/></svg>;
const IconMessageSquare = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>;
const IconSearch = ({ size = 24 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>;

const ToastMessage = ({ toast }) => {
  if (!toast) return null;
  const colors = {
    success: 'bg-green-100 text-green-800 border-green-300',
    error: 'bg-red-100 text-red-800 border-red-300',
    info: 'bg-blue-100 text-blue-800 border-blue-300',
    loading: 'bg-white text-slate-800 border-slate-300 shadow-xl'
  };
  return (
    <div className={`fixed bottom-8 right-4 left-4 md:left-auto md:right-8 px-4 sm:px-6 py-3 sm:py-4 rounded-lg shadow-lg border font-semibold flex items-center gap-3 z-[100] animate-bounce text-sm sm:text-base ${colors[toast.type]}`}>
      {toast.type === 'loading' ? (
        <div className="w-4 h-4 sm:w-5 sm:h-5 border-2 border-slate-400 border-t-transparent rounded-full animate-spin"></div>
      ) : toast.type === 'success' ? (
        <IconCheckCircle size={20} />
      ) : toast.type === 'error' ? (
        <div className="bg-red-200 text-red-800 rounded-full p-0.5"><IconPlus size={16} className="rotate-45" /></div>
      ) : (
        <IconCloud size={20} />
      )}
      <span>{toast.msg}</span>
    </div>
  );
};


// --- Constants ---
const PO_COLS = ['PO1', 'PO2', 'PO3', 'PO4', 'PO5', 'PO6', 'PO7', 'PO8', 'PO9', 'PO10', 'PO11', 'PO12'];
const PSO_COLS = ['PSO1', 'PSO2', 'PSO3'];
const ACHIEVING_SKILLS = [
  { key: 'knowledge', label: '1. Knowledge' }, { key: 'understanding', label: '2. Understanding' },
  { key: 'criticalSkills', label: '3. Critical Skills' }, { key: 'analyticalSkills', label: '4. Analytical Skills' },
  { key: 'problemSolving', label: '5. Problem Solving Skills' }, { key: 'practicalSkills', label: '6. Practical Skills' },
  { key: 'groupWork', label: '7. Group Work' }, { key: 'selfLearning', label: '8. Self-Learning' },
  { key: 'writtenComm', label: '9. Written Communication Skills' }, { key: 'verbalComm', label: '10. Verbal Communication Skills' },
  { key: 'presentation', label: '11. Presentation Skills' }, { key: 'behavioral', label: '12. Behavioral Skills' },
  { key: 'infoManagement', label: '13. Information Management' }, { key: 'personalManagement', label: '14. Personal Management' },
  { key: 'leadership', label: '15. Leadership Skills' }
];

const emptyFormState = {
  basicInfo: { courseTitle: '', courseCode: '', courseType: '', department: '', faculty: '' },
  summary: '',
  credits: { numberOfCredits: '', creditStructure: '', totalHours: '', weeksInSemester: '', departmentResponsible: '', totalCourseMarks: '', passCriterion: '', attendanceRequirement: '' },
  outcomes: [{ id: 'co1', text: '' }, { id: 'co2', text: '' }, { id: 'co3', text: '' }, { id: 'co4', text: '' }, { id: 'co5', text: '' }], 
  units: [{ id: 'u1', title: '', content: '' }],
  coPoMapping: {}, assessmentMapping: {}, achievingCos: {},
  teachingMethods: { faceToFace: '', demoVideos: '', demoModels: '', demoComputer: '', numeracySolving: '', pracCourseLab: '', pracCompLab: '', pracWorkshop: '', pracClinical: '', pracHospital: '', pracStudio: '', othersCaseStudy: '', othersGuest: '', othersIndustry: '', othersBrainstorming: '', othersGroup: '', othersInnovations: '', termTestsExams: '' },
  assessmentDetails: {
    para1: "The details of the components and subcomponents of course assessment are presented in the Programme Specifications document pertaining to the B. Tech. Programme. The procedure to determine the final course marks is also presented in the Programme Specifications document.",
    para2: "The evaluation questions are set to measure the attainment of the COs. In either component (CE or SEE) or subcomponent of CE (SC1, SC2, SC3 or SC4), COs are assessed as illustrated in the following Table.",
    para3: "The Course Leader assigned to the course, in consultation with the Head of the Department, shall provide the focus of course outcomes in each component assessed in the above template at the beginning of the semester.",
    para4: "Course reassessment policies are also presented in the Academic Regulations document."
  },
  resources: { essential: [], recommended: [], magazines: [], websites: [], electronic: [] },
  revisionComments: ''
};

// ==========================================
// 3. MICROSOFT GRAPH API SERVICES
// ==========================================
const getGraphToken = async () => {
  if (!msalInstance) throw new Error("MSAL not initialized");
  const account = msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("No active account!");
  
  try {
    const response = await msalInstance.acquireTokenSilent({ ...LOGIN_REQUEST, account });
    return response.accessToken;
  } catch  {
    const response = await msalInstance.acquireTokenPopup(LOGIN_REQUEST);
    return response.accessToken;
  }
};

const MSGraphService = {
  sharedFolderCache: null,

  async getTargetFolder(token) {
    if (this.sharedFolderCache) return this.sharedFolderCache;

    // STEP 1: Check if the logged-in user is the OWNER of the folder
    try {
      const myDriveRes = await fetch("https://graph.microsoft.com/v1.0/me/drive/root:/Course_Specifications", {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (myDriveRes.ok) {
        const data = await myDriveRes.json();
        this.sharedFolderCache = { driveId: data.parentReference.driveId, folderId: data.id };
        return this.sharedFolderCache;
      }
    } catch { /* Ignore and try step 2 */ }

    // STEP 2: Check if the folder was SHARED with the logged-in user
    const sharedRes = await fetch("https://graph.microsoft.com/v1.0/me/drive/sharedWithMe", {
      headers: { Authorization: `Bearer ${token}` }
    });
    
    if (sharedRes.ok) {
      const data = await sharedRes.json();
      const sharedFolder = data.value.find(item => item.name === "Course_Specifications");
      
      if (sharedFolder && sharedFolder.remoteItem) {
        this.sharedFolderCache = {
          driveId: sharedFolder.remoteItem.parentReference.driveId,
          folderId: sharedFolder.remoteItem.id
        };
        return this.sharedFolderCache;
      }
    }

    throw new Error("Folder 'Course_Specifications' not found. Ensure the owner explicitly shared it with this email address.");
  },

  async fetchAllSpecifications() {
    try {
      const token = await getGraphToken();
      const target = await this.getTargetFolder(token);
      
      const url = `https://graph.microsoft.com/v1.0/drives/${target.driveId}/items/${target.folderId}/children`;
      const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      
      if (response.status === 404) return []; 
      if (!response.ok) throw new Error("Failed to fetch files from the repository.");

      const data = await response.json();
      const files = data.value.filter(f => f.name.endsWith('.json'));
      
      const docs = await Promise.all(files.map(async f => {
          const content = await this.getSpecificationContent(f.id);
          return {
              id: f.id,
              name: content?.basicInfo?.courseCode || f.name.replace('.json', ''),
              courseTitle: content?.basicInfo?.courseTitle || 'Untitled Course',
              author: content?.author || f.createdBy?.user?.displayName || 'Unknown',
              lastModified: new Date(f.lastModifiedDateTime).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
              status: content?.status || 'Draft',
              data: content
          };
      }));
      return docs;
    } catch (e) { 
      console.error(e); 
      throw e; 
    }
  },

  async getSpecificationContent(fileId) {
    try {
      const token = await getGraphToken();
      const target = await this.getTargetFolder(token);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${target.driveId}/items/${fileId}/content`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      return await response.json();
    } catch (e) { 
      console.error("Error reading file:", e); 
      return null; 
    }
  },

  async saveSpecification(courseCode, formData, newStatus, accountName, originalAuthor) {
    try {
      const token = await getGraphToken();
      const target = await this.getTargetFolder(token);
      
      const safeCode = courseCode.replace(/[^a-zA-Z0-9]/g, '_') || `Draft_${Date.now()}`;
      const fileName = `${safeCode}.json`;
      
      const finalAuthor = originalAuthor || accountName;
      const payload = { ...formData, status: newStatus, author: finalAuthor };
      
      const url = `https://graph.microsoft.com/v1.0/drives/${target.driveId}/items/${target.folderId}:/${fileName}:/content`;
      
      const response = await fetch(url, {
        method: 'PUT',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      
      if (!response.ok) throw new Error("Access Denied: Could not save to repository.");
      return true;
    } catch (e) {
      console.error("Failed to save JSON:", e);
      throw e;
    }
  },

  async uploadDocxBlob(courseCode, blob) {
    try {
      const token = await getGraphToken();
      const target = await this.getTargetFolder(token);
      
      const safeCode = courseCode.replace(/[^a-zA-Z0-9]/g, '_') || `Draft_${Date.now()}`;
      const fileName = `${safeCode}_Approved.docx`;
      
      // Save in a subfolder "Approved_Docs" to avoid clutter
      const url = `https://graph.microsoft.com/v1.0/drives/${target.driveId}/items/${target.folderId}:/Approved_Docs/${fileName}:/content`;
      
      const response = await fetch(url, {
        method: 'PUT',
        headers: { 
          Authorization: `Bearer ${token}`, 
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
        },
        body: blob
      });
      
      if (!response.ok) throw new Error("Failed to upload Word Document.");
      return true;
    } catch (e) {
      console.error("Failed to upload DOCX:", e);
      throw e;
    }
  }
};
// ==========================================
// 4. DOCX RENDERING ENGINE
// ==========================================

// Utilities to calculate durations dynamically
const getTeachingTotals = (tm) => {
  const parse = (v) => parseInt(v) || 0;
  const demo = parse(tm?.demoVideos) + parse(tm?.demoModels) + parse(tm?.demoComputer);
  const num = parse(tm?.numeracySolving);
  const prac = parse(tm?.pracCourseLab) + parse(tm?.pracCompLab) + parse(tm?.pracWorkshop) + parse(tm?.pracClinical) + parse(tm?.pracHospital) + parse(tm?.pracStudio);
  const others = parse(tm?.othersCaseStudy) + parse(tm?.othersGuest) + parse(tm?.othersIndustry) + parse(tm?.othersBrainstorming) + parse(tm?.othersGroup) + parse(tm?.othersInnovations);
  const faceToFace = parse(tm?.faceToFace);
  const termTests = parse(tm?.termTestsExams);
  
  return {
    demo, num, prac, others, faceToFace, termTests,
    total: demo + num + prac + others + faceToFace + termTests
  };
};

const generateDocumentContent = (d) => {
  const t = getTeachingTotals(d.teachingMethods);
  
  return `
      <h1 class="header-title">M S Ramaiah University of Applied Sciences</h1>
      <h2 class="header-subtitle">Course Specifications: ${d.basicInfo?.courseTitle || 'Untitled Course'}</h2>
      
      <table>
          <tr><td width="25%" class="bold bg-gray">Course Title</td><td width="25%">${d.basicInfo?.courseTitle || ''}</td><td width="25%" class="bold bg-gray">Course Code</td><td width="25%">${d.basicInfo?.courseCode || ''}</td></tr>
          <tr><td class="bold bg-gray">Course Type</td><td>${d.basicInfo?.courseType || ''}</td><td class="bold bg-gray">Department</td><td>${d.basicInfo?.department || ''}</td></tr>
          <tr><td class="bold bg-gray">Faculty</td><td colspan="3">${d.basicInfo?.faculty || ''}</td></tr>
      </table>

      <h3>1. Course Summary</h3>
      <div class="section-text"><p>${d.summary || ''}</p></div>

      <h3>2. Course Size and Credits</h3>
      <table>
          <tr><td width="25%" class="bold bg-gray">Number of Credits</td><td width="25%">${d.credits?.numberOfCredits || ''}</td><td width="25%" class="bold bg-gray">Credit Structure (L:T:P)</td><td width="25%">${d.credits?.creditStructure || ''}</td></tr>
          <tr><td class="bold bg-gray">Total Hours of Interaction</td><td>${d.credits?.totalHours || ''}</td><td class="bold bg-gray">Weeks in a Semester</td><td>${d.credits?.weeksInSemester || ''}</td></tr>
          <tr><td class="bold bg-gray">Department Responsible</td><td colspan="3">${d.credits?.departmentResponsible || ''}</td></tr>
          <tr><td class="bold bg-gray">Total Course Marks</td><td>${d.credits?.totalCourseMarks || ''}</td><td class="bold bg-gray">Pass Criterion</td><td>${d.credits?.passCriterion || ''}</td></tr>
          <tr><td class="bold bg-gray">Attendance Requirement</td><td colspan="3">${d.credits?.attendanceRequirement || ''}</td></tr>
      </table>

      <h3>3. Course Outcomes (COs)</h3>
      <div class="section-text">
        <p>After the successful completion of this course, the student will be able to:</p>
        <ul style="list-style-type: none; padding-left: 0;">
            ${d.outcomes?.map((co, i) => `<li><span class="bold">CO-${i+1}.</span> ${co.text}</li>`).join('') || ''}
        </ul>
      </div>

      <h3>4. Course Contents</h3>
      <div class="section-text">
        ${d.units?.map((u, i) => `<p><span class="bold">Unit ${i+1} (${u.title || 'Untitled'}):</span> ${u.content || ''}</p>`).join('') || ''}
      </div>

      <h3>5. CO-PO Mapping</h3>
      <table>
          <tr>
              <th rowspan="2" class="center" width="8%">COs</th>
              <th colspan="12" class="center">Programme Outcomes (POs)</th>
              <th colspan="3" class="center">PSOs</th>
          </tr>
          <tr>
              ${PO_COLS.map(po => `<th class="center">${po.replace('PO', 'PO-')}</th>`).join('')}
              ${PSO_COLS.map(pso => `<th class="center">${pso.replace('PSO', 'PSO-')}</th>`).join('')}
          </tr>
          ${d.outcomes?.map((co, i) => `
          <tr>
              <td class="center bold bg-gray">CO-${i+1}</td>
              ${PO_COLS.map(po => `<td class="center">${d.coPoMapping?.[co.id]?.[po] || ''}</td>`).join('')}
              ${PSO_COLS.map(pso => `<td class="center">${d.coPoMapping?.[co.id]?.[pso] || ''}</td>`).join('')}
          </tr>`).join('') || ''}
          <tr><td colspan="16" class="center" style="font-size: 9pt;"><i>3: Very Strong Contribution, 2: Strong Contribution, 1: Moderate Contribution</i></td></tr>
      </table>

      <h3>6. Course Teaching and Learning Methods</h3>
      <table>
          <tr><th width="65%">Teaching and Learning Methods</th><th class="center" width="20%">Duration in hours</th><th class="center" width="15%">Total Duration in Hours</th></tr>
          
          <tr><td colspan="2" class="bold">Face to Face Lectures</td><td class="center">${t.faceToFace}</td></tr>
          
          <tr><td colspan="2" class="bold">Demonstrations</td><td rowspan="4" class="center align-middle">${t.demo}</td></tr>
          <tr><td>1. Demonstration using Videos</td><td class="center">${d.teachingMethods?.demoVideos || '00'}</td></tr>
          <tr><td>2. Demonstration using Physical Models / Systems</td><td class="center">${d.teachingMethods?.demoModels || '00'}</td></tr>
          <tr><td>3. Demonstration on a Computer</td><td class="center">${d.teachingMethods?.demoComputer || '00'}</td></tr>
          
          <tr><td colspan="2" class="bold">Numeracy</td><td rowspan="2" class="center align-middle">${t.num}</td></tr>
          <tr><td>1. Solving Numerical Problems</td><td class="center">${d.teachingMethods?.numeracySolving || '00'}</td></tr>
          
          <tr><td colspan="2" class="bold">Practical Work</td><td rowspan="7" class="center align-middle">${t.prac}</td></tr>
          <tr><td>1. Course Laboratory</td><td class="center">${d.teachingMethods?.pracCourseLab || '00'}</td></tr>
          <tr><td>2. Computer Laboratory</td><td class="center">${d.teachingMethods?.pracCompLab || '00'}</td></tr>
          <tr><td>3. Engineering Workshop / Course/Workshop / Kitchen</td><td class="center">${d.teachingMethods?.pracWorkshop || '00'}</td></tr>
          <tr><td>4. Clinical Laboratory</td><td class="center">${d.teachingMethods?.pracClinical || '00'}</td></tr>
          <tr><td>5. Hospital</td><td class="center">${d.teachingMethods?.pracHospital || '00'}</td></tr>
          <tr><td>6. Model Studio</td><td class="center">${d.teachingMethods?.pracStudio || '00'}</td></tr>
          
          <tr><td colspan="2" class="bold">Others</td><td rowspan="7" class="center align-middle">${t.others}</td></tr>
          <tr><td>1. Case Study Presentation</td><td class="center">${d.teachingMethods?.othersCaseStudy || '00'}</td></tr>
          <tr><td>2. Guest Lecture</td><td class="center">${d.teachingMethods?.othersGuest || '00'}</td></tr>
          <tr><td>3. Industry / Field Visit</td><td class="center">${d.teachingMethods?.othersIndustry || '00'}</td></tr>
          <tr><td>4. Brain Storming Sessions</td><td class="center">${d.teachingMethods?.othersBrainstorming || '00'}</td></tr>
          <tr><td>5. Group Discussions</td><td class="center">${d.teachingMethods?.othersGroup || '00'}</td></tr>
          <tr><td>6. Discussing Possible Innovations</td><td class="center">${d.teachingMethods?.othersInnovations || '00'}</td></tr>
          
          <tr><td colspan="2" class="bold">Term Tests, Laboratory Examination/Written Examination, Presentations</td><td class="center align-middle">${t.termTests}</td></tr>
          <tr><td colspan="2" class="bold right-align">Total Duration in Hours:</td><td class="center bold">${t.total}</td></tr>
      </table>

      <h3>7. Course Assessment and Reassessment</h3>
      <div class="section-text">
        <p>${d.assessmentDetails?.para1 || ''}</p>
        <p>${d.assessmentDetails?.para2 || ''}</p>
      </div>
      
      <table>
          <tr>
              <th colspan="4" class="center">Focus of COs on each Component or Subcomponent of Evaluation</th>
          </tr>
          <tr>
              <th width="25%"></th>
              <th colspan="2" class="center">Component 1: CE (50% Weightage)</th>
              <th width="25%" class="center">Component 2: SEE (50% Weightage)</th>
          </tr>
          <tr>
              <th class="bold right-align">Subcomponent ►</th>
              <th></th>
              <th></th>
              <th rowspan="3" class="center align-middle">100 Marks</th>
          </tr>
          <tr>
              <th class="bold right-align">Subcomponent Type ►</th>
              <th class="center">Term Tests</th>
              <th class="center">Assignments</th>
          </tr>
          <tr>
              <th class="bold right-align">Maximum Marks ►</th>
              <th class="center">50</th>
              <th class="center">50</th>
          </tr>
          ${d.outcomes?.map((co, i) => `
          <tr>
              <td class="center bold bg-gray">CO-${i+1}</td>
              <td class="center">${d.assessmentMapping?.[co.id]?.termTests ? 'X' : ''}</td>
              <td class="center">${d.assessmentMapping?.[co.id]?.assignments ? 'X' : ''}</td>
              <td class="center">${d.assessmentMapping?.[co.id]?.see ? 'X' : ''}</td>
          </tr>
          `).join('') || ''}
          <tr><td colspan="4" class="center" style="font-size: 9pt;">The details of number of tests and assignments to be conducted are presented in the Academic Regulations and Programme Specifications Document.</td></tr>
      </table>
      
      <div class="section-text">
        <p>${d.assessmentDetails?.para3 || ''}</p>
        <p>${d.assessmentDetails?.para4 || ''}</p>
      </div>

      <h3>8. Achieving Course Learning Outcomes</h3>
      <table>
          <tr><th width="10%" class="center">S. No</th><th width="45%">Curriculum and Capabilities Skills</th><th width="45%">How imparted during the course</th></tr>
          ${ACHIEVING_SKILLS.map((s, i) => `<tr><td class="center">${i+1}</td><td class="bold">${s.label.replace(/^\d+\.\s*/, '')}</td><td>${d.achievingCos?.[s.key] || '--'}</td></tr>`).join('')}
      </table>

      <h3>9. Course Resources</h3>
      <div class="section-text">
        <p class="bold">a. Essential Reading</p>
        <ol>${d.resources?.essential?.length ? d.resources.essential.map(r => `<li>${r.text}</li>`).join('') : '<li>None</li>'}</ol>
        
        <p class="bold" style="margin-top: 12px;">b. Recommended Reading</p>
        <ol>${d.resources?.recommended?.length ? d.resources.recommended.map(r => `<li>${r.text}</li>`).join('') : '<li>None</li>'}</ol>
        
        <p class="bold" style="margin-top: 12px;">c. Magazines and Journals</p>
        <ol>${d.resources?.magazines?.length ? d.resources.magazines.map(r => `<li>${r.text}</li>`).join('') : '<li>None</li>'}</ol>
        
        <p class="bold" style="margin-top: 12px;">d. Websites</p>
        <ol>${d.resources?.websites?.length ? d.resources.websites.map(r => `<li>${r.text}</li>`).join('') : '<li>None</li>'}</ol>
        
        <p class="bold" style="margin-top: 12px;">e. Other Electronic Resources</p>
        <ol>${d.resources?.electronic?.length ? d.resources.electronic.map(r => `<li>${r.text}</li>`).join('') : '<li>None</li>'}</ol>
      </div>
  `;
};

const getDocumentCSS = () => `
  body { font-family: 'Times New Roman', Times, serif; font-size: 11pt; color: #000; line-height: 1.4; margin: 0; padding: 0; position: relative; }
  .header-title { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 5px; color: #000; text-transform: uppercase; }
  .header-subtitle { text-align: center; font-size: 12pt; margin-bottom: 25px; font-weight: bold; color: #000; }
  h3 { font-size: 11.5pt; font-weight: bold; margin-top: 25px; margin-bottom: 12px; color: #000; padding-bottom: 2px; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 20px; table-layout: fixed; }
  th, td { border: 1px solid #000; padding: 5px 6px; text-align: left; vertical-align: top; word-wrap: break-word; font-size: 10pt; color: #000; }
  th { font-weight: bold; color: #000; }
  ul, ol { margin-top: 6px; margin-bottom: 6px; padding-left: 24px; }
  li { margin-bottom: 4px; }
  .center { text-align: center; }
  .right-align { text-align: right; }
  .align-middle { vertical-align: middle; }
  .bold { font-weight: bold; }
  .bg-gray { background-color: #f2f2f2; }
  .section-text { padding: 0 5px; margin-bottom: 10px; }
  .content-wrapper { position: relative; z-index: 1; }
`;

// Extract Docx Engine Generation Core
const DocxEngine = {
  async generateBlob(d) {
    const loadDocxJS = async () => {
      if (window.docx) return window.docx;
      const loadScript = (src) => new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = () => {
            if (window.docx) resolve(window.docx);
            else reject(new Error("Script loaded but docx object is missing"));
        };
        script.onerror = () => reject(new Error("Failed to load script from " + src));
        document.head.appendChild(script);
      });
      try {
        return await loadScript("https://cdn.jsdelivr.net/npm/docx@7.8.2/build/index.js");
      } catch {
        return await loadScript("https://unpkg.com/docx@7.8.2/build/index.js");
      }
    };

    const docx = await loadDocxJS();

    const cell = (text, opts = {}) => {
        const cellOptions = {
            children: [new docx.Paragraph({
                children: [new docx.TextRun({ text: String(text || ""), bold: opts.bold, size: 20 })],
                alignment: opts.align === 'center' ? docx.AlignmentType.CENTER : opts.align === 'right' ? docx.AlignmentType.RIGHT : docx.AlignmentType.LEFT,
            })],
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            verticalAlign: opts.vAlign === 'middle' ? docx.VerticalAlign.CENTER : docx.VerticalAlign.TOP
        };
        if (opts.colSpan) cellOptions.columnSpan = opts.colSpan;
        if (opts.rowSpan) cellOptions.rowSpan = opts.rowSpan;
        if (opts.bg) cellOptions.shading = { fill: opts.bg };
        if (opts.width) cellOptions.width = { size: opts.width, type: docx.WidthType.PERCENTAGE };
        return new docx.TableCell(cellOptions);
    };

    const heading = (text, lvl) => {
        const hOpts = {
            children: [new docx.TextRun({ text, bold: true, size: lvl === docx.HeadingLevel.HEADING_1 ? 28 : lvl === docx.HeadingLevel.HEADING_2 ? 24 : 22, color: "000000" })],
            heading: lvl,
            alignment: lvl === docx.HeadingLevel.HEADING_1 || lvl === docx.HeadingLevel.HEADING_2 ? docx.AlignmentType.CENTER : docx.AlignmentType.LEFT,
            spacing: { before: 240, after: 120 }
        };
        return new docx.Paragraph(hOpts);
    };

    const p = (text) => new docx.Paragraph({ children: [new docx.TextRun({ text: text || "", size: 22, color: "000000" })], spacing: { after: 120 } });
    
    const tableOpts = { width: { size: 100, type: docx.WidthType.PERCENTAGE } };
    const t = getTeachingTotals(d.teachingMethods);

    // 3. Course Outcomes (COs) List
    const coParas = [new docx.Paragraph({ text: "After the successful completion of this course, the student will be able to:", spacing: { after: 120 } })];
    d.outcomes?.forEach((co, i) => {
        coParas.push(new docx.Paragraph({
            children: [
                new docx.TextRun({ text: `CO-${i+1}. `, bold: true, size: 22 }),
                new docx.TextRun({ text: co.text || '', size: 22 })
            ],
            spacing: { after: 80 }
        }));
    });

    // 4. Course Contents List
    const unitParas = d.units?.map((u, i) => {
        return new docx.Paragraph({
            children: [
                new docx.TextRun({ text: `Unit ${i+1} (${u.title || 'Untitled'}): `, bold: true, size: 22 }),
                new docx.TextRun({ text: u.content || '', size: 22 })
            ],
            spacing: { after: 120 }
        });
    }) || [];

    // 5. CO-PO Mapping Array Setup
    const coPoRows = [];
    coPoRows.push(new docx.TableRow({
        children: [
            cell("COs", { rowSpan: 2, bold: true, align: 'center', width: 10 }),
            cell("Programme Outcomes (POs)", { colSpan: 12, bold: true, align: 'center', width: 75 }),
            cell("PSOs", { colSpan: 3, bold: true, align: 'center', width: 15 })
        ]
    }));
    coPoRows.push(new docx.TableRow({
        children: [
            ...PO_COLS.map(po => cell(po.replace('PO', 'PO-'), { bold: true, align: 'center' })),
            ...PSO_COLS.map(pso => cell(pso.replace('PSO', 'PSO-'), { bold: true, align: 'center' }))
        ]
    }));
    d.outcomes?.forEach((co, i) => {
        coPoRows.push(new docx.TableRow({
            children: [
                cell(`CO-${i+1}`, { bold: true, align: 'center', bg: "F2F2F2" }),
                ...PO_COLS.map(po => cell(d.coPoMapping?.[co.id]?.[po] || '', { align: 'center' })),
                ...PSO_COLS.map(pso => cell(d.coPoMapping?.[co.id]?.[pso] || '', { align: 'center' }))
            ]
        }));
    });
    coPoRows.push(new docx.TableRow({ children: [ cell("3: Very Strong Contribution, 2: Strong Contribution, 1: Moderate Contribution", { colSpan: 16, align: 'center' }) ] }));

    // 6. Teaching Methods Array Setup
    const teachingRows = [
        new docx.TableRow({ children: [cell("Teaching and Learning Methods", { bold: true, width: 65 }), cell("Duration in hours", { bold: true, align: 'center', width: 20 }), cell("Total Duration in Hours", { bold: true, align: 'center', width: 15 })] }),
        new docx.TableRow({ children: [cell("Face to Face Lectures", { bold: true, colSpan: 2 }), cell(String(t.faceToFace), { align: 'center' })] }),
        
        new docx.TableRow({ children: [cell("Demonstrations", { bold: true, colSpan: 2 }), cell(String(t.demo), { rowSpan: 4, align: 'center', vAlign: 'middle' })] }),
        new docx.TableRow({ children: [cell("1. Demonstration using Videos"), cell(d.teachingMethods?.demoVideos || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("2. Demonstration using Physical Models / Systems"), cell(d.teachingMethods?.demoModels || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("3. Demonstration on a Computer"), cell(d.teachingMethods?.demoComputer || '00', { align: 'center' })] }),
        
        new docx.TableRow({ children: [cell("Numeracy", { bold: true, colSpan: 2 }), cell(String(t.num), { rowSpan: 2, align: 'center', vAlign: 'middle' })] }),
        new docx.TableRow({ children: [cell("1. Solving Numerical Problems"), cell(d.teachingMethods?.numeracySolving || '00', { align: 'center' })] }),
        
        new docx.TableRow({ children: [cell("Practical Work", { bold: true, colSpan: 2 }), cell(String(t.prac), { rowSpan: 7, align: 'center', vAlign: 'middle' })] }),
        new docx.TableRow({ children: [cell("1. Course Laboratory"), cell(d.teachingMethods?.pracCourseLab || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("2. Computer Laboratory"), cell(d.teachingMethods?.pracCompLab || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("3. Engineering Workshop / Course/Workshop / Kitchen"), cell(d.teachingMethods?.pracWorkshop || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("4. Clinical Laboratory"), cell(d.teachingMethods?.pracClinical || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("5. Hospital"), cell(d.teachingMethods?.pracHospital || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("6. Model Studio"), cell(d.teachingMethods?.pracStudio || '00', { align: 'center' })] }),
        
        new docx.TableRow({ children: [cell("Others", { bold: true, colSpan: 2 }), cell(String(t.others), { rowSpan: 7, align: 'center', vAlign: 'middle' })] }),
        new docx.TableRow({ children: [cell("1. Case Study Presentation"), cell(d.teachingMethods?.othersCaseStudy || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("2. Guest Lecture"), cell(d.teachingMethods?.othersGuest || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("3. Industry / Field Visit"), cell(d.teachingMethods?.othersIndustry || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("4. Brain Storming Sessions"), cell(d.teachingMethods?.othersBrainstorming || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("5. Group Discussions"), cell(d.teachingMethods?.othersGroup || '00', { align: 'center' })] }),
        new docx.TableRow({ children: [cell("6. Discussing Possible Innovations"), cell(d.teachingMethods?.othersInnovations || '00', { align: 'center' })] }),
        
        new docx.TableRow({ children: [cell("Term Tests, Laboratory Examination/Written Examination, Presentations", { bold: true, colSpan: 2 }), cell(String(t.termTests), { align: 'center' })] }),
        new docx.TableRow({ children: [cell("Total Duration in Hours:", { bold: true, align: 'right', colSpan: 2 }), cell(String(t.total), { bold: true, align: 'center' })] }),
    ];

    // 7. Assessment Mapping Array Setup
    const assessmentRows = [
        new docx.TableRow({ children: [ cell("Focus of COs on each Component or Subcomponent of Evaluation", { colSpan: 4, bold: true, align: 'center' }) ] }),
        new docx.TableRow({ children: [
            cell("", { width: 25 }),
            cell("Component 1: CE (50% Weightage)", { colSpan: 2, bold: true, align: 'center', width: 50 }),
            cell("Component 2: SEE (50% Weightage)", { bold: true, align: 'center', width: 25 })
        ]}),
        new docx.TableRow({ children: [
            cell("Subcomponent ►", { bold: true, align: 'right' }),
            cell(""),
            cell(""),
            cell("100 Marks", { rowSpan: 3, bold: true, align: 'center', vAlign: 'middle' })
        ]}),
        new docx.TableRow({ children: [
            cell("Subcomponent Type ►", { bold: true, align: 'right' }),
            cell("Term Tests", { bold: true, align: 'center' }),
            cell("Assignments", { bold: true, align: 'center' })
        ]}),
        new docx.TableRow({ children: [
            cell("Maximum Marks ►", { bold: true, align: 'right' }),
            cell("50", { bold: true, align: 'center' }),
            cell("50", { bold: true, align: 'center' })
        ]})
    ];
    d.outcomes?.forEach((co, i) => {
        assessmentRows.push(new docx.TableRow({
            children: [
                cell(`CO-${i+1}`, { bold: true, align: 'center', bg: "F2F2F2" }),
                cell(d.assessmentMapping?.[co.id]?.termTests ? 'X' : '', { align: 'center' }),
                cell(d.assessmentMapping?.[co.id]?.assignments ? 'X' : '', { align: 'center' }),
                cell(d.assessmentMapping?.[co.id]?.see ? 'X' : '', { align: 'center' }),
            ]
        }));
    });
    assessmentRows.push(new docx.TableRow({ children: [ cell("The details of number of tests and assignments to be conducted are presented in the Academic Regulations and Programme Specifications Document.", { colSpan: 4, align: 'center' }) ] }));

    // 8. Achieving Rows Setup
    const achievingRows = [
        new docx.TableRow({
            children: [
                cell("S. No", { bold: true, align: 'center', width: 10 }),
                cell("Curriculum and Capabilities Skills", { bold: true, width: 45 }),
                cell("How imparted during the course", { bold: true, width: 45 })
            ]
        }),
        ...ACHIEVING_SKILLS.map((s, i) => new docx.TableRow({
            children: [
                cell(String(i+1), { align: 'center' }),
                cell(s.label.replace(/^\d+\.\s*/, ''), { bold: true }),
                cell(d.achievingCos?.[s.key] || '--')
            ]
        }))
    ];

    const renderList = (items) => {
        if (!items || items.length === 0) return [new docx.Paragraph({ text: "None", bullet: { level: 0 } })];
        return items.map(item => new docx.Paragraph({ text: item.text, bullet: { level: 0 } }));
    };

    const doc = new docx.Document({
        creator: "Course Spec Portal",
        sections: [{
            properties: {},
            children: [
                heading("M S Ramaiah University of Applied Sciences", docx.HeadingLevel.HEADING_1),
                heading(`Course Specifications: ${d.basicInfo?.courseTitle || 'Untitled Course'}`, docx.HeadingLevel.HEADING_2),
                new docx.Table({
                    ...tableOpts,
                    rows: [
                        new docx.TableRow({ children: [cell("Course Title", { bold: true, bg: "F2F2F2", width: 25 }), cell(d.basicInfo?.courseTitle, { width: 25 }), cell("Course Code", { bold: true, bg: "F2F2F2", width: 25 }), cell(d.basicInfo?.courseCode, { width: 25 })] }),
                        new docx.TableRow({ children: [cell("Course Type", { bold: true, bg: "F2F2F2" }), cell(d.basicInfo?.courseType), cell("Department", { bold: true, bg: "F2F2F2" }), cell(d.basicInfo?.department)] }),
                        new docx.TableRow({ children: [cell("Faculty", { bold: true, bg: "F2F2F2" }), cell(d.basicInfo?.faculty, { colSpan: 3 })] })
                    ]
                }),
                heading("1. Course Summary", docx.HeadingLevel.HEADING_3), p(d.summary),
                heading("2. Course Size and Credits", docx.HeadingLevel.HEADING_3),
                new docx.Table({
                    ...tableOpts,
                    rows: [
                        new docx.TableRow({ children: [cell("Number of Credits", { bold: true, bg: "F2F2F2", width: 25 }), cell(d.credits?.numberOfCredits, { width: 25 }), cell("Credit Structure (L:T:P)", { bold: true, bg: "F2F2F2", width: 25 }), cell(d.credits?.creditStructure, { width: 25 })] }),
                        new docx.TableRow({ children: [cell("Total Hours of Interaction", { bold: true, bg: "F2F2F2" }), cell(d.credits?.totalHours), cell("Weeks in a Semester", { bold: true, bg: "F2F2F2" }), cell(d.credits?.weeksInSemester)] }),
                        new docx.TableRow({ children: [cell("Department Responsible", { bold: true, bg: "F2F2F2" }), cell(d.credits?.departmentResponsible, { colSpan: 3 })] }),
                        new docx.TableRow({ children: [cell("Total Course Marks", { bold: true, bg: "F2F2F2" }), cell(d.credits?.totalCourseMarks), cell("Pass Criterion", { bold: true, bg: "F2F2F2" }), cell(d.credits?.passCriterion)] }),
                        new docx.TableRow({ children: [cell("Attendance Requirement", { bold: true, bg: "F2F2F2" }), cell(d.credits?.attendanceRequirement, { colSpan: 3 })] })
                    ]
                }),
                heading("3. Course Outcomes (COs)", docx.HeadingLevel.HEADING_3), ...coParas,
                heading("4. Course Contents", docx.HeadingLevel.HEADING_3), ...unitParas,
                heading("5. CO-PO Mapping", docx.HeadingLevel.HEADING_3), new docx.Table({ ...tableOpts, rows: coPoRows }),
                heading("6. Course Teaching and Learning Methods", docx.HeadingLevel.HEADING_3), new docx.Table({ ...tableOpts, rows: teachingRows }),
                heading("7. Course Assessment and Reassessment", docx.HeadingLevel.HEADING_3),
                p(d.assessmentDetails?.para1), p(d.assessmentDetails?.para2),
                new docx.Table({ ...tableOpts, rows: assessmentRows }),
                new docx.Paragraph({ text: "" }),
                p(d.assessmentDetails?.para3), p(d.assessmentDetails?.para4),
                heading("8. Achieving Course Learning Outcomes", docx.HeadingLevel.HEADING_3), new docx.Table({ ...tableOpts, rows: achievingRows }),
                heading("9. Course Resources", docx.HeadingLevel.HEADING_3),
                new docx.Paragraph({ children: [new docx.TextRun({ text: "a. Essential Reading", bold: true, size: 22 })], spacing: { before: 120, after: 60 } }), ...renderList(d.resources?.essential),
                new docx.Paragraph({ children: [new docx.TextRun({ text: "b. Recommended Reading", bold: true, size: 22 })], spacing: { before: 120, after: 60 } }), ...renderList(d.resources?.recommended),
                new docx.Paragraph({ children: [new docx.TextRun({ text: "c. Magazines and Journals", bold: true, size: 22 })], spacing: { before: 120, after: 60 } }), ...renderList(d.resources?.magazines),
                new docx.Paragraph({ children: [new docx.TextRun({ text: "d. Websites", bold: true, size: 22 })], spacing: { before: 120, after: 60 } }), ...renderList(d.resources?.websites),
                new docx.Paragraph({ children: [new docx.TextRun({ text: "e. Other Electronic Resources", bold: true, size: 22 })], spacing: { before: 120, after: 60 } }), ...renderList(d.resources?.electronic),
            ]
        }]
    });

    return await docx.Packer.toBlob(doc);
  },

  async download(d, showToast) {
    try {
      const blob = await this.generateBlob(d);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${d.basicInfo?.courseCode || 'Course'}_Specifications.docx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(url), 100);
      if(showToast) showToast("Document exported successfully!", "success");
    } catch (error) {
      console.error("Native Generator Failed:", error);
      if(showToast) showToast("Exporting using basic fallback format due to network issues.", "info");
      
      const oldHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head><meta charset='utf-8'><title>Course Specifications</title><style>${getDocumentCSS()}</style></head>
        <body><div class="content-wrapper">${generateDocumentContent(d)}</div></body></html>`;
      const blob = new Blob(['\ufeff', oldHtml], { type: 'application/msword' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${d.basicInfo?.courseCode || 'Course'}_Specifications.doc`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  }
};


// ==========================================
// 5. IN-APP IMMERSIVE VIEWER
// ==========================================
const DocumentViewer = ({ docMeta, docData, onBack, onEdit, canEdit, showToast, role, onStatusChange }) => {
  const [isExporting, setIsExporting] = useState(false);
  const [showRevisionModal, setShowRevisionModal] = useState(false);
  const [revisionCommentsInput, setRevisionCommentsInput] = useState('');

  const handleExport = async () => {
    setIsExporting(true);
    showToast("Generating Word Document...", "loading");
    await DocxEngine.download(docData, showToast);
    setIsExporting(false);
  };

  const handleHODAction = async (status) => {
    let comments = null;
    if (status === 'Needs Revision') {
       comments = revisionCommentsInput;
    }
    const updatedData = { ...docData };
    if (comments !== null) {
      updatedData.revisionComments = comments;
    }
    await onStatusChange(updatedData, status);
  };

  return (
    <div className="bg-[#f3f2f1] min-h-screen flex flex-col font-sans relative">
      {/* Mobile-Friendly App Bar */}
      <div className="bg-[#1e1e1e] text-white px-4 md:px-5 py-3 flex flex-col sm:flex-row items-start sm:items-center justify-between shadow-md z-20 sticky top-0 gap-4 sm:gap-0">
        <div className="flex items-center gap-3 w-full sm:w-auto">
          <button onClick={onBack} className="p-2 hover:bg-white/10 rounded-full transition shrink-0"><IconArrowLeft size={20} /></button>
          <div className="flex items-center gap-3 truncate w-full">
            <div className="bg-white/10 p-2 rounded-lg text-white shrink-0"><IconFileText size={20}/></div>
            <div className="truncate">
              <h2 className="font-bold text-sm leading-tight tracking-wide truncate">{docData?.basicInfo?.courseTitle || 'Untitled Document'}</h2>
              <p className="text-xs text-gray-400 truncate">Status: <span className="text-white">{docMeta?.status}</span> • Last modified by {docMeta?.author}</p>
            </div>
          </div>
        </div>
        <div className="flex flex-wrap gap-2 w-full sm:w-auto">
          {role === 'HOD' && docMeta?.status !== 'Approved' && (
            <>
              <button onClick={() => setShowRevisionModal(true)} className="flex-1 sm:flex-none bg-yellow-500 text-white hover:bg-yellow-600 px-4 py-2 rounded-md text-sm font-bold shadow transition flex items-center justify-center gap-2">
                Revise
              </button>
              <button onClick={() => handleHODAction('Approved')} className="flex-1 sm:flex-none bg-green-600 text-white hover:bg-green-700 px-4 py-2 rounded-md text-sm font-bold shadow transition flex items-center justify-center gap-2">
                <IconCheckCircle size={16}/> Approve
              </button>
            </>
          )}
          {canEdit && <button onClick={onEdit} className="flex-1 sm:flex-none bg-white/10 hover:bg-white/20 text-white px-4 py-2 rounded-md text-sm font-semibold transition flex items-center justify-center gap-2"><IconEdit size={16}/> Edit</button>}
          <button onClick={handleExport} disabled={isExporting} className="flex-1 sm:flex-none bg-[#0f6cbd] text-white hover:bg-[#0c5697] px-4 py-2 rounded-md text-sm font-bold shadow transition flex items-center justify-center gap-2 disabled:opacity-50">
            <IconDownload size={16}/> {isExporting ? 'Generating...' : 'Export'}
          </button>
        </div>
      </div>
      
      {/* Revision Modal for HOD */}
      {showRevisionModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[100] p-4 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-xl p-5 md:p-6 w-full max-w-lg shadow-2xl">
            <h3 className="text-xl font-bold text-gray-900 mb-2 flex items-center gap-2">
              <IconMessageSquare size={22} className="text-yellow-600"/> Request Revision
            </h3>
            <p className="text-sm text-gray-500 mb-4">Provide detailed feedback to the faculty on required changes.</p>
            <textarea 
              className="w-full p-3 md:p-4 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#0f6cbd] outline-none resize-y mb-6 text-sm" 
              rows="5" placeholder="e.g., Update CO-PO mapping..."
              value={revisionCommentsInput} onChange={e => setRevisionCommentsInput(e.target.value)}
            />
            <div className="flex flex-col sm:flex-row justify-end gap-3">
              <button onClick={() => setShowRevisionModal(false)} className="w-full sm:w-auto px-5 py-2 text-gray-600 hover:bg-gray-100 font-semibold rounded-md transition border border-gray-200 sm:border-transparent">Cancel</button>
              <button onClick={() => { handleHODAction('Needs Revision'); setShowRevisionModal(false); }} className="w-full sm:w-auto px-5 py-2 bg-yellow-500 text-white rounded-md font-bold hover:bg-yellow-600 shadow transition flex justify-center items-center gap-2">
                Send Revision Request
              </button>
            </div>
          </div>
        </div>
      )}

      {/* The Viewer Canvas */}
      <div className="flex-1 overflow-auto p-2 sm:p-8 flex flex-col items-center pb-20 bg-gray-200 w-full">
        {/* We use min-w-[600px] and let the parent scroll to act like a real document reader on mobile */}
        <div className="bg-white shadow-2xl rounded-sm w-full max-w-[850px] min-w-[300px] overflow-x-auto relative">
            
           {/* Document Logo Watermark */}
           <div className="fixed inset-0 flex justify-center items-center pointer-events-none overflow-hidden z-0 min-w-[600px]">
             <img src={logo} alt="Watermark" className="w-[80%] max-w-[100px] opacity-[0.07] object-contain" />
           </div>

           {docData?.revisionComments && docMeta?.status === 'Needs Revision' && (
             <div className="bg-yellow-50 border-b border-yellow-200 p-4 m-1 rounded-t-sm shadow-sm relative z-10 min-w-[600px]">
               <div className="flex items-start gap-3">
                 <div className="text-yellow-600 mt-0.5"><IconMessageSquare size={20}/></div>
                 <div>
                   <h4 className="text-yellow-800 font-bold text-sm">Revision Requested by HOD</h4>
                   <p className="text-yellow-700 text-sm mt-1 whitespace-pre-wrap">{docData.revisionComments}</p>
                 </div>
               </div>
             </div>
           )}
           
           <div className="p-6 sm:p-16 pt-8 sm:pt-12 relative z-10 bg-transparent min-w-[600px]">
             <style>{getDocumentCSS()}</style>
             <div className="content-wrapper" dangerouslySetInnerHTML={{ __html: generateDocumentContent(docData) }} />
           </div>
        </div>
      </div>
    </div>
  );
};

// ==========================================
// 6. FORM EDITOR COMPONENT
// ==========================================
const CourseForm = ({ initialData, currentStatus, role, onSave, onCancel }) => {
  const [formData, setFormData] = useState(() => {
      const data = initialData ? JSON.parse(JSON.stringify(initialData)) : JSON.parse(JSON.stringify(emptyFormState));
      if (!data.assessmentDetails) {
          data.assessmentDetails = emptyFormState.assessmentDetails;
      }
      return data;
  });

  const [isSaving, setIsSaving] = useState(false);

  const handleNestedChange = (cat, f, val) => setFormData(p => ({ ...p, [cat]: { ...p[cat], [f]: val } }));
  const handleTextChange = (f, val) => setFormData(p => ({ ...p, [f]: val }));
  const addDynamicItem = (cat) => setFormData(p => ({ ...p, [cat]: [...p[cat], { id: `item_${Date.now()}`, text: '' }] }));
  const removeDynamicItem = (cat, id) => setFormData(p => ({ ...p, [cat]: p[cat].filter(i => i.id !== id) }));
  const handleDynamicListChange = (cat, id, val) => setFormData(p => ({ ...p, [cat]: p[cat].map(i => i.id === id ? { ...i, text: val } : i) }));
  
  const addUnit = () => setFormData(p => ({ ...p, units: [...p.units, { id: Date.now(), title: '', content: '' }] }));
  const updateUnit = (id, f, val) => setFormData(p => ({ ...p, units: p.units.map(u => u.id === id ? { ...u, [f]: val } : u) }));
  const removeUnit = (id) => setFormData(p => ({ ...p, units: p.units.filter(u => u.id !== id) }));
  
  const handleCoPoChange = (coId, colKey, val) => setFormData(p => ({ ...p, coPoMapping: { ...p.coPoMapping, [coId]: { ...(p.coPoMapping[coId] || {}), [colKey]: val } } }));
  const handleAssessmentChange = (coId, type, checked) => setFormData(p => ({ ...p, assessmentMapping: { ...p.assessmentMapping, [coId]: { ...(p.assessmentMapping[coId] || {}), [type]: checked } } }));

  const addResource = (t) => setFormData(p => ({ ...p, resources: { ...p.resources, [t]: [...p.resources[t], { id: Date.now(), text: '' }] } }));
  const updateResource = (t, id, val) => setFormData(p => ({ ...p, resources: { ...p.resources, [t]: p.resources[t].map(r => r.id === id ? { ...r, text: val } : r) } }));
  const removeResource = (t, id) => setFormData(p => ({ ...p, resources: { ...p.resources, [t]: p.resources[t].filter(r => r.id !== id) } }));

  const handleSubmit = async (newStatus) => {
    setIsSaving(true);
    await onSave({ ...formData }, newStatus);
    setIsSaving(false);
  };

  const teachingTotals = getTeachingTotals(formData.teachingMethods);

  return (
    <div className="bg-white shadow-2xl rounded-xl overflow-hidden mt-2 md:mt-6 mb-12 max-w-[1000px] mx-auto relative border border-gray-200">
      
      {/* Mobile-Friendly Sticky Form Header */}
      <div className="bg-white border-b border-gray-200 px-4 md:px-8 py-3 md:py-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-3 sticky top-0 z-30 shadow-sm">
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={onCancel} className="p-2 bg-gray-50 border border-gray-200 rounded-full hover:bg-gray-100 text-gray-600 transition shrink-0"><IconArrowLeft size={18} /></button>
          <div className="truncate flex-1">
              <h2 className="text-lg md:text-xl font-bold text-gray-900 tracking-tight truncate">Course Specification Editor</h2>
              <p className="text-xs text-gray-500 flex items-center gap-1"><IconCloud size={12} className="shrink-0"/> <span className="truncate">Changes save to OneDrive folder</span></p>
          </div>
        </div>
        
        {/* Buttons Array (Stacks or Wraps nicely) */}
        <div className="flex flex-wrap gap-2 w-full md:w-auto">
           {role === 'FACULTY' && (
             <>
                <button onClick={() => handleSubmit('Draft')} disabled={isSaving} className="flex-1 md:flex-none bg-white border border-gray-300 text-gray-700 px-4 py-2 rounded-md text-sm md:text-base font-semibold hover:bg-gray-50 transition shadow-sm disabled:opacity-50">
                  Save Draft
                </button>
                <button onClick={() => handleSubmit('Submitted')} disabled={isSaving} className="flex-1 md:flex-none bg-[#0f6cbd] text-white px-4 py-2 rounded-md text-sm md:text-base font-bold shadow-md hover:bg-[#0c5697] transition flex justify-center items-center gap-2 disabled:opacity-50">
                   Submit for Review
                </button>
             </>
           )}
           {role === 'HOD' && (
             <button onClick={() => handleSubmit(currentStatus)} disabled={isSaving} className="flex-1 md:flex-none bg-[#0f6cbd] text-white px-4 py-2 rounded-md text-sm md:text-base font-bold shadow-md hover:bg-[#0c5697] transition flex justify-center items-center gap-2 disabled:opacity-50">
               Save Changes
             </button>
           )}
        </div>
      </div>

      {formData.revisionComments && (
        <div className={`mx-4 md:mx-8 mt-4 md:mt-6 p-4 rounded-lg shadow-sm border ${currentStatus === 'Needs Revision' ? 'bg-yellow-50 border-yellow-200' : 'bg-slate-50 border-slate-200'}`}>
          <div className="flex items-start gap-3">
            <div className={`${currentStatus === 'Needs Revision' ? 'text-yellow-600' : 'text-slate-500'} mt-0.5 shrink-0`}><IconMessageSquare size={20}/></div>
            <div>
              <h4 className={`${currentStatus === 'Needs Revision' ? 'text-yellow-800' : 'text-slate-700'} font-bold text-sm`}>
                {currentStatus === 'Needs Revision' ? 'Revision Requested by HOD' : 'Previous Review Comments'}
              </h4>
              <p className={`${currentStatus === 'Needs Revision' ? 'text-yellow-700' : 'text-slate-600'} text-sm mt-1 whitespace-pre-wrap`}>{formData.revisionComments}</p>
            </div>
          </div>
        </div>
      )}

      <div className="p-4 sm:p-8 space-y-8 md:space-y-12">
        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconSettings size={20} className="text-gray-400"/> Course Identification</h2>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 md:gap-6 bg-gray-50 p-4 md:p-6 rounded-lg border border-gray-100">
            <div><label className="block text-xs font-bold text-gray-600 uppercase mb-1.5">Course Title <span className="text-red-500">*</span></label><input type="text" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none text-sm md:text-base" placeholder="e.g., Data Structures" value={formData.basicInfo.courseTitle} onChange={(e) => handleNestedChange('basicInfo', 'courseTitle', e.target.value)} /></div>
            <div><label className="block text-xs font-bold text-gray-600 uppercase mb-1.5">Course Code <span className="text-red-500">*</span></label><input type="text" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none font-mono uppercase text-sm md:text-base" placeholder="e.g., CSD201A" value={formData.basicInfo.courseCode} onChange={(e) => handleNestedChange('basicInfo', 'courseCode', e.target.value.toUpperCase())} /></div>
            <div><label className="block text-xs font-bold text-gray-600 uppercase mb-1.5">Course Type</label><input type="text" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none text-sm md:text-base" placeholder="e.g., Core Theory" value={formData.basicInfo.courseType} onChange={(e) => handleNestedChange('basicInfo', 'courseType', e.target.value)} /></div>
            <div className="lg:col-span-2"><label className="block text-xs font-bold text-gray-600 uppercase mb-1.5">Department</label><input type="text" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none text-sm md:text-base" placeholder="e.g., Computer Science" value={formData.basicInfo.department} onChange={(e) => handleNestedChange('basicInfo', 'department', e.target.value)} /></div>
            <div><label className="block text-xs font-bold text-gray-600 uppercase mb-1.5">Faculty</label><input type="text" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none text-sm md:text-base" placeholder="e.g., Eng. & Tech." value={formData.basicInfo.faculty} onChange={(e) => handleNestedChange('basicInfo', 'faculty', e.target.value)} /></div>
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconBookOpen size={20} className="text-gray-400"/> 1. Course Summary</h2>
          <textarea rows="4" className="w-full p-3 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none resize-y text-sm" placeholder="Enter an academic summary..." value={formData.summary} onChange={(e) => handleTextChange('summary', e.target.value)} />
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconTarget size={20} className="text-gray-400"/> 2. Course Size and Credits</h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 md:gap-5 bg-gray-50 p-4 md:p-6 rounded-lg border border-gray-100">
            {[
              { label: 'Credits', key: 'numberOfCredits', placeholder: '03' },
              { label: 'Structure (L:T:P)', key: 'creditStructure', placeholder: '3:0:0' },
              { label: 'Total Hours', key: 'totalHours', placeholder: '45' },
              { label: 'Weeks/Semester', key: 'weeksInSemester', placeholder: '15' },
              { label: 'Dept Responsible', key: 'departmentResponsible', placeholder: 'CSE' },
              { label: 'Total Marks', key: 'totalCourseMarks', placeholder: '100' },
              { label: 'Pass Criterion', key: 'passCriterion', placeholder: 'As per Reg.' },
              { label: 'Attendance Req.', key: 'attendanceRequirement', placeholder: 'As per Reg.' }
            ].map((f) => (
              <div key={f.key} className={f.key === 'departmentResponsible' ? 'md:col-span-2' : ''}>
                <label className="block text-xs font-bold text-gray-500 uppercase mb-1.5">{f.label}</label>
                <input type="text" className="w-full p-2 text-sm border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none bg-white" placeholder={f.placeholder} value={formData.credits[f.key]} onChange={(e) => handleNestedChange('credits', f.key, e.target.value)} />
              </div>
            ))}
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center justify-between">
            <div className="flex items-center gap-2"><IconCheckCircle size={20} className="text-gray-400"/> <span className="hidden sm:inline">3. Course Outcomes (COs)</span><span className="sm:hidden">3. COs</span></div>
            <button type="button" onClick={() => addDynamicItem('outcomes')} className="text-xs sm:text-sm bg-gray-100 text-gray-700 font-semibold px-2 sm:px-3 py-1.5 rounded-md flex items-center gap-1 hover:bg-gray-200"><IconPlus size={16} /> Add <span className="hidden sm:inline">CO</span></button>
          </h2>
          <div className="space-y-3">
            {formData.outcomes.map((co, index) => (
              <div key={co.id} className="flex gap-2 sm:gap-3 items-start group">
                <span className="font-bold text-gray-600 mt-2 sm:mt-2.5 w-10 sm:w-12 text-xs sm:text-sm shrink-0">CO-{index + 1}.</span>
                <input type="text" className="flex-1 w-full p-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none text-sm" placeholder="e.g., Describe Data Structures..." value={co.text} onChange={(e) => handleDynamicListChange('outcomes', co.id, e.target.value)} />
                <button type="button" onClick={() => removeDynamicItem('outcomes', co.id)} className="p-2 text-red-400 hover:text-red-600 rounded-md bg-red-50 mt-0.5 opacity-100 sm:opacity-0 group-hover:opacity-100 transition shrink-0"><IconTrash2 size={18} /></button>
              </div>
            ))}
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center justify-between">
            <div className="flex items-center gap-2"><IconList size={20} className="text-gray-400"/> <span className="hidden sm:inline">4. Course Contents (Units)</span><span className="sm:hidden">4. Units</span></div>
            <button type="button" onClick={addUnit} className="text-xs sm:text-sm bg-gray-100 text-gray-700 font-semibold px-2 sm:px-3 py-1.5 rounded-md flex items-center gap-1 hover:bg-gray-200"><IconPlus size={16} /> Add Unit</button>
          </h2>
          <div className="space-y-4">
            {formData.units.map((unit, index) => (
              <div key={unit.id} className="bg-white border border-gray-200 rounded-lg p-4 md:p-6 relative group shadow-sm">
                <button type="button" onClick={() => removeUnit(unit.id)} className="absolute top-2 md:top-4 right-2 md:right-4 text-red-500 hover:text-red-700 bg-red-50 p-1.5 rounded-md opacity-100 md:opacity-0 group-hover:opacity-100 transition"><IconTrash2 size={16} /></button>
                <div className="mb-4 pr-8 md:pr-10">
                  <label className="block text-xs font-bold text-gray-500 uppercase mb-1.5">Unit {index + 1} Title</label>
                  <input type="text" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none text-sm md:text-base" placeholder="e.g., Stacks and Queues" value={unit.title} onChange={(e) => updateUnit(unit.id, 'title', e.target.value)} />
                </div>
                <div>
                  <label className="block text-xs font-bold text-gray-500 uppercase mb-1.5">Topics Covered</label>
                  <textarea rows="3" className="w-full p-2.5 border border-gray-300 rounded-md focus:ring-2 focus:ring-[#0f6cbd] outline-none resize-y text-sm" placeholder="List the contents of this unit..." value={unit.content} onChange={(e) => updateUnit(unit.id, 'content', e.target.value)} />
                </div>
              </div>
            ))}
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconGrid size={20} className="text-gray-400"/> 5. CO-PO Mapping</h2>
          <p className="text-xs text-gray-500 mb-2 sm:hidden italic">Swipe horizontally to view full table</p>
          <div className="overflow-x-auto border border-gray-300 rounded-lg shadow-sm w-full">
            <table className="min-w-full divide-y divide-gray-200 text-sm table-fixed sm:table-auto w-max sm:w-full">
              <thead className="bg-gray-100">
                <tr>
                  <th rowSpan="2" className="px-2 sm:px-3 py-2 text-left text-xs font-bold text-gray-700 border-r border-b w-12 sticky left-0 bg-gray-100">COs</th>
                  <th colSpan="12" className="px-2 sm:px-3 py-2 text-center text-xs font-bold text-gray-700 border-r border-b">Programme Outcomes (POs)</th>
                  <th colSpan="3" className="px-2 sm:px-3 py-2 text-center text-xs font-bold text-gray-700 border-b">PSOs</th>
                </tr>
                <tr>
                  {PO_COLS.map(po => <th key={po} className="px-1 py-1.5 text-center text-[9px] sm:text-[10px] font-bold text-gray-600 border-r border-b w-8">{po.replace('PO', 'PO-')}</th>)}
                  {PSO_COLS.map(pso => <th key={pso} className="px-1 py-1.5 text-center text-[9px] sm:text-[10px] font-bold text-gray-600 border-r border-b w-8">{pso.replace('PSO', 'PSO-')}</th>)}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {formData.outcomes.map((co, index) => (
                  <tr key={co.id} className="hover:bg-gray-50">
                    <td className="px-2 sm:px-3 py-1.5 whitespace-nowrap text-xs font-bold text-gray-900 border-r bg-gray-50 sticky left-0 z-10 shadow-[1px_0_0_rgba(0,0,0,0.1)]">CO-{index + 1}</td>
                    {[...PO_COLS, ...PSO_COLS].map(colKey => (
                      <td key={colKey} className="p-0.5 border-r text-center w-8">
                        <input type="text" maxLength="1" className="w-full text-center p-1 border border-transparent hover:border-gray-300 focus:border-[#0f6cbd] rounded outline-none text-xs font-bold text-gray-800 transition-colors" value={(formData.coPoMapping[co.id] && formData.coPoMapping[co.id][colKey]) || ''} onChange={(e) => handleCoPoChange(co.id, colKey, e.target.value)} />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-gray-500 mt-2 text-center italic">3: Very Strong Contribution, 2: Strong Contribution, 1: Moderate Contribution</p>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconActivity size={20} className="text-gray-400"/> 6. Teaching and Learning Methods</h2>
          <p className="text-xs text-gray-500 mb-2 sm:hidden italic">Swipe horizontally to view full table</p>
          <div className="border border-gray-300 rounded-lg shadow-sm overflow-x-auto w-full bg-white">
            <table className="min-w-full divide-y divide-gray-200 text-sm table-fixed sm:table-auto w-max sm:w-full">
              <thead className="bg-gray-100">
                <tr>
                  <th className="px-3 sm:px-4 py-2 sm:py-3 text-left text-[10px] sm:text-xs font-bold text-gray-700 uppercase border-r w-[250px] sm:w-3/5">Method</th>
                  <th className="px-2 sm:px-4 py-2 sm:py-3 text-center text-[10px] sm:text-xs font-bold text-gray-700 uppercase border-r w-24 sm:w-1/5">Duration (Hours)</th>
                  <th className="px-2 sm:px-4 py-2 sm:py-3 text-center text-[10px] sm:text-xs font-bold text-gray-700 uppercase w-24 sm:w-1/5">Total Group Hours</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {/* Face to Face */}
                <tr>
                   <td className="px-3 sm:px-4 py-2 font-bold text-gray-800 border-r whitespace-normal">Face to Face Lectures</td>
                   <td className="px-2 sm:px-4 py-2 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border border-gray-300 rounded focus:ring-2 focus:ring-[#0f6cbd] outline-none" value={formData.teachingMethods.faceToFace} onChange={(e) => handleNestedChange('teachingMethods', 'faceToFace', e.target.value)} /></td>
                   <td className="px-2 sm:px-4 py-2 text-center font-bold text-gray-900 bg-gray-50">{teachingTotals.faceToFace}</td>
                </tr>
                {/* Demonstrations */}
                <tr className="bg-gray-50"><td colSpan="2" className="px-3 sm:px-4 py-2 font-bold text-gray-800 border-r">Demonstrations</td><td rowSpan="4" className="text-center align-middle font-bold text-gray-900 text-base sm:text-lg">{teachingTotals.demo}</td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">1. Demonstration using Videos</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.demoVideos} onChange={(e) => handleNestedChange('teachingMethods', 'demoVideos', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">2. Demonstration using Physical Models / Systems</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.demoModels} onChange={(e) => handleNestedChange('teachingMethods', 'demoModels', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal border-b">3. Demonstration on a Computer</td><td className="px-2 py-1.5 text-center border-r border-b"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.demoComputer} onChange={(e) => handleNestedChange('teachingMethods', 'demoComputer', e.target.value)} /></td></tr>
                {/* Numeracy */}
                <tr className="bg-gray-50"><td colSpan="2" className="px-3 sm:px-4 py-2 font-bold text-gray-800 border-r">Numeracy</td><td rowSpan="2" className="text-center align-middle font-bold text-gray-900 text-base sm:text-lg">{teachingTotals.num}</td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal border-b">1. Solving Numerical Problems</td><td className="px-2 py-1.5 text-center border-r border-b"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.numeracySolving} onChange={(e) => handleNestedChange('teachingMethods', 'numeracySolving', e.target.value)} /></td></tr>
                {/* Practical Work */}
                <tr className="bg-gray-50"><td colSpan="2" className="px-3 sm:px-4 py-2 font-bold text-gray-800 border-r">Practical Work</td><td rowSpan="7" className="text-center align-middle font-bold text-gray-900 text-base sm:text-lg">{teachingTotals.prac}</td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">1. Course Laboratory</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.pracCourseLab} onChange={(e) => handleNestedChange('teachingMethods', 'pracCourseLab', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">2. Computer Laboratory</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.pracCompLab} onChange={(e) => handleNestedChange('teachingMethods', 'pracCompLab', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">3. Engineering Workshop / Kitchen</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.pracWorkshop} onChange={(e) => handleNestedChange('teachingMethods', 'pracWorkshop', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">4. Clinical Laboratory</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.pracClinical} onChange={(e) => handleNestedChange('teachingMethods', 'pracClinical', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">5. Hospital</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.pracHospital} onChange={(e) => handleNestedChange('teachingMethods', 'pracHospital', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal border-b">6. Model Studio</td><td className="px-2 py-1.5 text-center border-r border-b"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.pracStudio} onChange={(e) => handleNestedChange('teachingMethods', 'pracStudio', e.target.value)} /></td></tr>
                {/* Others */}
                <tr className="bg-gray-50"><td colSpan="2" className="px-3 sm:px-4 py-2 font-bold text-gray-800 border-r">Others</td><td rowSpan="7" className="text-center align-middle font-bold text-gray-900 text-base sm:text-lg">{teachingTotals.others}</td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">1. Case Study Presentation</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.othersCaseStudy} onChange={(e) => handleNestedChange('teachingMethods', 'othersCaseStudy', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">2. Guest Lecture</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.othersGuest} onChange={(e) => handleNestedChange('teachingMethods', 'othersGuest', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">3. Industry / Field Visit</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.othersIndustry} onChange={(e) => handleNestedChange('teachingMethods', 'othersIndustry', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">4. Brain Storming Sessions</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.othersBrainstorming} onChange={(e) => handleNestedChange('teachingMethods', 'othersBrainstorming', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal">5. Group Discussions</td><td className="px-2 py-1.5 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.othersGroup} onChange={(e) => handleNestedChange('teachingMethods', 'othersGroup', e.target.value)} /></td></tr>
                <tr><td className="pl-6 sm:pl-8 pr-2 sm:pr-4 py-1.5 border-r text-gray-600 text-xs whitespace-normal border-b">6. Discussing Possible Innovations</td><td className="px-2 py-1.5 text-center border-r border-b"><input type="number" className="w-12 sm:w-16 text-center p-1 border rounded text-xs focus:ring-2 outline-none" value={formData.teachingMethods.othersInnovations} onChange={(e) => handleNestedChange('teachingMethods', 'othersInnovations', e.target.value)} /></td></tr>
                
                {/* Term Tests */}
                <tr>
                   <td className="px-3 sm:px-4 py-2 font-bold text-gray-800 border-r whitespace-normal text-xs sm:text-sm">Term Tests, Laboratory Examination/Written Examination, Presentations</td>
                   <td className="px-2 sm:px-4 py-2 text-center border-r"><input type="number" className="w-12 sm:w-16 text-center p-1 border border-gray-300 rounded focus:ring-2 focus:ring-[#0f6cbd] outline-none" value={formData.teachingMethods.termTestsExams} onChange={(e) => handleNestedChange('teachingMethods', 'termTestsExams', e.target.value)} /></td>
                   <td className="px-2 sm:px-4 py-2 text-center font-bold text-gray-900 bg-gray-50">{teachingTotals.termTests}</td>
                </tr>

                {/* Total Row */}
                <tr className="bg-blue-50">
                  <td colSpan="2" className="px-3 sm:px-4 py-3 text-right font-bold text-[#0f6cbd] uppercase tracking-wider text-[10px] sm:text-xs border-r">Total Duration in Hours</td>
                  <td className="px-2 sm:px-4 py-3 text-center font-black text-[#0f6cbd] text-lg sm:text-xl">{teachingTotals.total}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconTarget size={20} className="text-gray-400"/> 7. Assessment and Reassessment</h2>
          
          <div className="space-y-3 mb-6">
            <label className="block text-[10px] sm:text-xs font-bold text-gray-500 uppercase">Assessment Details Paragraphs (Editable)</label>
            <textarea className="w-full p-2.5 border border-gray-300 rounded text-sm outline-none focus:ring-2 focus:ring-[#0f6cbd]" rows="3" value={formData.assessmentDetails?.para1} onChange={e => handleNestedChange('assessmentDetails', 'para1', e.target.value)} />
            <textarea className="w-full p-2.5 border border-gray-300 rounded text-sm outline-none focus:ring-2 focus:ring-[#0f6cbd]" rows="3" value={formData.assessmentDetails?.para2} onChange={e => handleNestedChange('assessmentDetails', 'para2', e.target.value)} />
          </div>

          <p className="text-xs text-gray-500 mb-2 sm:hidden italic">Swipe horizontally to view full table</p>
          <div className="overflow-x-auto border border-gray-300 rounded-lg shadow-sm w-full">
            <table className="min-w-full divide-y divide-gray-200 text-xs sm:text-sm table-fixed sm:table-auto w-max sm:w-full">
              <thead className="bg-gray-100">
                <tr>
                  <th colSpan="4" className="px-2 sm:px-4 py-2 sm:py-3 text-center font-bold text-gray-800 border-b text-[10px] sm:text-sm whitespace-normal">Focus of COs on each Component or Subcomponent of Evaluation</th>
                </tr>
                <tr>
                  <th width="20%" className="px-2 sm:px-4 py-2 border-r border-b"></th>
                  <th colSpan="2" className="px-2 sm:px-4 py-2 text-center font-bold text-gray-800 border-r border-b">Component 1: CE (50%)</th>
                  <th width="30%" className="px-2 sm:px-4 py-2 text-center font-bold text-gray-800 border-b">Component 2: SEE (50%)</th>
                </tr>
                <tr>
                  <th className="px-2 sm:px-4 py-2 text-right font-bold text-gray-800 border-r border-b">Subcomponent ►</th>
                  <th className="px-2 sm:px-4 py-2 border-r border-b"></th>
                  <th className="px-2 sm:px-4 py-2 border-r border-b"></th>
                  <th rowSpan="3" className="px-2 sm:px-4 py-2 text-center align-middle font-bold text-gray-800 border-b">100 Marks</th>
                </tr>
                <tr>
                  <th className="px-2 sm:px-4 py-2 text-right font-bold text-gray-800 border-r border-b">Type ►</th>
                  <th className="px-2 sm:px-4 py-2 text-center font-bold text-gray-700 border-r border-b">Term Tests</th>
                  <th className="px-2 sm:px-4 py-2 text-center font-bold text-gray-700 border-r border-b">Assignments</th>
                </tr>
                <tr>
                  <th className="px-2 sm:px-4 py-2 text-right font-bold text-gray-800 border-r border-b">Max Marks ►</th>
                  <th className="px-2 sm:px-4 py-2 text-center font-bold text-gray-700 border-r border-b">50</th>
                  <th className="px-2 sm:px-4 py-2 text-center font-bold text-gray-700 border-r border-b">50</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {formData.outcomes.map((co, index) => (
                  <tr key={co.id} className="hover:bg-gray-50">
                    <td className="px-2 sm:px-4 py-2 text-center font-bold text-gray-800 border-r bg-gray-50">CO-{index + 1}</td>
                    <td className="px-2 sm:px-4 py-2 text-center border-r"><input type="checkbox" className="w-4 h-4 sm:w-5 sm:h-5 text-[#0f6cbd] rounded cursor-pointer" checked={formData.assessmentMapping[co.id]?.termTests || false} onChange={(e) => handleAssessmentChange(co.id, 'termTests', e.target.checked)} /></td>
                    <td className="px-2 sm:px-4 py-2 text-center border-r"><input type="checkbox" className="w-4 h-4 sm:w-5 sm:h-5 text-[#0f6cbd] rounded cursor-pointer" checked={formData.assessmentMapping[co.id]?.assignments || false} onChange={(e) => handleAssessmentChange(co.id, 'assignments', e.target.checked)} /></td>
                    <td className="px-2 sm:px-4 py-2 text-center"><input type="checkbox" className="w-4 h-4 sm:w-5 sm:h-5 text-[#0f6cbd] rounded cursor-pointer" checked={formData.assessmentMapping[co.id]?.see || false} onChange={(e) => handleAssessmentChange(co.id, 'see', e.target.checked)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          
          <div className="space-y-3 mt-6">
            <textarea className="w-full p-2.5 border border-gray-300 rounded text-sm outline-none focus:ring-2 focus:ring-[#0f6cbd]" rows="3" value={formData.assessmentDetails?.para3} onChange={e => handleNestedChange('assessmentDetails', 'para3', e.target.value)} />
            <textarea className="w-full p-2.5 border border-gray-300 rounded text-sm outline-none focus:ring-2 focus:ring-[#0f6cbd]" rows="2" value={formData.assessmentDetails?.para4} onChange={e => handleNestedChange('assessmentDetails', 'para4', e.target.value)} />
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconActivity size={20} className="text-gray-400"/> 8. Achieving Course Learning Outcomes</h2>
          <p className="text-xs text-gray-500 mb-2 sm:hidden italic">Swipe horizontally to view full table</p>
          <div className="border border-gray-300 rounded-lg shadow-sm overflow-x-auto w-full">
            <table className="min-w-full divide-y divide-gray-200 text-xs sm:text-sm table-fixed sm:table-auto w-max sm:w-full">
              <thead className="bg-gray-100">
                <tr>
                  <th className="px-3 sm:px-4 py-2 sm:py-3 text-left font-bold text-gray-700 uppercase border-r w-40 sm:w-1/3">Skills</th>
                  <th className="px-3 sm:px-4 py-2 sm:py-3 text-left font-bold text-gray-700 uppercase w-48 sm:w-2/3">How imparted during the course</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {ACHIEVING_SKILLS.map((skill) => (
                  <tr key={skill.key} className="hover:bg-gray-50">
                    <td className="px-3 sm:px-4 py-2 font-semibold text-gray-800 border-r whitespace-normal">{skill.label}</td>
                    <td className="px-2 py-1"><input type="text" className="w-full p-2 text-sm border border-transparent hover:border-gray-300 focus:border-[#0f6cbd] rounded outline-none transition bg-transparent focus:bg-white" placeholder="e.g., Classroom lectures" value={formData.achievingCos[skill.key] || ''} onChange={(e) => handleNestedChange('achievingCos', skill.key, e.target.value)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <section>
          <h2 className="text-lg font-bold text-gray-900 border-b pb-2 mb-4 flex items-center gap-2"><IconBookOpen size={20} className="text-gray-400"/> 9. Course Resources</h2>
          <div className="space-y-6">
            {[
              { key: 'essential', label: 'a. Essential Reading' },
              { key: 'recommended', label: 'b. Recommended Reading' },
              { key: 'magazines', label: 'c. Magazines and Journals' },
              { key: 'websites', label: 'd. Websites' },
              { key: 'electronic', label: 'e. Other Electronic Resources' }
            ].map(category => (
              <div key={category.key} className="bg-gray-50 p-4 md:p-5 rounded-lg border border-gray-100">
                <div className="flex justify-between items-center mb-3">
                  <label className="block text-sm font-bold text-gray-800">{category.label}</label>
                  <button type="button" onClick={() => addResource(category.key)} className="text-xs bg-white border border-gray-300 text-gray-700 font-semibold px-2.5 py-1.5 rounded-md flex items-center gap-1 hover:bg-gray-100 shadow-sm transition">
                    <IconPlus size={14} /> <span className="hidden sm:inline">Add Resource</span><span className="sm:hidden">Add</span>
                  </button>
                </div>
                <div className="space-y-2">
                  {formData.resources[category.key].length === 0 && <p className="text-xs text-gray-400 italic ml-2">No resources added.</p>}
                  {formData.resources[category.key].map((item, idx) => (
                    <div key={item.id} className="flex gap-2 items-start group bg-white p-1 sm:p-1.5 rounded shadow-sm border border-gray-100">
                      <span className="text-gray-400 mt-2 text-xs sm:text-sm font-mono w-6 text-center shrink-0">{idx + 1}.</span>
                      <input type="text" className="flex-1 w-full p-2 text-xs sm:text-sm border-none focus:ring-0 outline-none" placeholder="Enter resource details..." value={item.text} onChange={(e) => updateResource(category.key, item.id, e.target.value)} />
                      <button type="button" onClick={() => removeResource(category.key, item.id)} className="p-2 text-red-400 hover:text-red-600 rounded opacity-100 sm:opacity-0 group-hover:opacity-100 transition shrink-0"><IconTrash2 size={16} /></button>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </section>
      </div>
    </div>
  );
};


// ==========================================
// 7. MAIN DASHBOARD APPLICATION
// ==========================================
export default function DashboardApp() {
  const [view, setView] = useState(() => {
    const hash = window.location.hash.replace('#', '').toUpperCase();
    return ['LOGIN', 'DASHBOARD', 'FORM', 'VIEWER'].includes(hash) ? hash : 'LOGIN';
  }); 
  const [role, setRole] = useState(null); 
  const [account, setAccount] = useState(null);
  const [isMsalReady, setIsMsalReady] = useState(false);
  
  const navigate = useCallback((newView) => {
    setView(newView);
    window.location.hash = newView;
  }, []);
  
  const [documents, setDocuments] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedDocMeta, setSelectedDocMeta] = useState(null);
  const [selectedDocData, setSelectedDocData] = useState(null);
  
  // Custom Toast State
  const [toast, setToast] = useState(null);

  const showToast = useCallback((msg, type = 'info') => {
    setToast({ msg, type });
    if (type !== 'loading') {
      setTimeout(() => setToast(null), 4000);
    }
  }, []);

  useEffect(() => {
    const handlePopState = () => {
      const hash = window.location.hash.replace('#', '').toUpperCase();
      // We rely entirely on the strict render guard for security now.
      // This listener only needs to sync the URL hash with the view state.
      setView(['LOGIN', 'DASHBOARD', 'FORM', 'VIEWER'].includes(hash) ? hash : 'DASHBOARD');
    };
    
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []); // <-- Empty dependency array prevents the stale closure bug

  const refreshDocuments = useCallback(async (silent = false) => {
    if (!silent) showToast("Syncing with OneDrive...", "loading");
    try {
      if (MSAL_CONFIG.auth.clientId !== "YOUR_MS_CLIENT_ID_HERE") {
          const files = await MSGraphService.fetchAllSpecifications();
          setDocuments(files);
      }
      if (!silent) setToast(null); 
    } catch (err) {
      if (!silent) showToast(err.message || "Failed to fetch files from OneDrive.", "error");
      else console.error("Silent auto-sync failed:", err);
    }
  }, [showToast]); 

  useEffect(() => {
    const loadMsal = () => new Promise((resolve, reject) => {
      if (window.msal) return resolve();
      const script = document.createElement('script');
      script.src = "https://alcdn.msauth.net/browser/2.37.1/js/msal-browser.min.js";
      script.async = true;
      script.onload = resolve;
      script.onerror = reject;
      document.head.appendChild(script);
    });

    loadMsal().then(() => {
      msalInstance = new window.msal.PublicClientApplication(MSAL_CONFIG);
      const activeAccount = msalInstance.getAllAccounts()[0];
      if (activeAccount) {
        setAccount(activeAccount);
        setRole(determineUserRole(activeAccount));
        navigate('DASHBOARD');
        refreshDocuments();
      }
      setIsMsalReady(true);
    }).catch(e => {
      console.error("Failed to load MSAL script", e);
      showToast("Identity service blocked or offline. Check network.", "error");
    });
  }, [navigate, refreshDocuments, showToast]);

  // Auto-refresh polling effect
  useEffect(() => {
    let intervalId;
    if (account && view === 'DASHBOARD') {
      // Poll quietly every 15 seconds to check for document changes
      intervalId = setInterval(() => {
        refreshDocuments(true);
      }, 15000);
    }
    return () => {
      if (intervalId) clearInterval(intervalId);
    };
  }, [account, view, refreshDocuments]);

  const handleMicrosoftLogin = async () => {
    if (!isMsalReady) return;
    try {
      const response = await msalInstance.loginPopup(LOGIN_REQUEST);
      
      // Tell MSAL to set this as the active session
      msalInstance.setActiveAccount(response.account);
      
      setAccount(response.account);
      setRole(determineUserRole(response.account));
      
      // Navigate will now smoothly transition to the dashboard without fighting the event listener
      navigate('DASHBOARD');
      showToast(`Welcome, ${response.account.name}!`, "success");
      refreshDocuments();
    } catch (e) { 
      console.error("Login failed", e);
      showToast(`Login Canceled or Failed. Check console.`, "error");
    }
  };

  const handleLogout = () => {
    // Local App Logout instead of complete Microsoft logout
    sessionStorage.clear(); // Clears MSAL cache tokens locally
    setAccount(null);
    setRole(null);
    
    // Replace the history state so the user cannot click "Back" to return to the dashboard
    window.history.replaceState(null, '', '#LOGIN');
    setView('LOGIN');
    
    showToast("Successfully signed out.", "info");
  };

  const openNewForm = () => {
    setSelectedDocMeta({ status: 'Draft', author: account?.name });
    setSelectedDocData(null);
    navigate('FORM');
  };

  const openDocumentViewer = async (docMeta) => {
    showToast("Opening document...", "loading");
    try {
      let data = emptyFormState;
      if (MSAL_CONFIG.auth.clientId !== "YOUR_MS_CLIENT_ID_HERE") {
          data = await MSGraphService.getSpecificationContent(docMeta.id);
      } else {
          data = docMeta.data || emptyFormState;
      }
      setSelectedDocMeta(docMeta);
      setSelectedDocData(data || emptyFormState);
      setToast(null);
      navigate('VIEWER');
    } catch {
      showToast("Could not load document content.", "error");
    }
  };

  const saveDocument = async (formData, newStatus) => {
    const code = formData.basicInfo.courseCode;
    if (!code) {
      showToast("Course Code is mandatory to save.", "error");
      return;
    }
    
    showToast("Saving changes...", "loading");
    const originalAuthor = selectedDocMeta?.author;
    
    try {
      if (MSAL_CONFIG.auth.clientId !== "YOUR_MS_CLIENT_ID_HERE") {
        // Step 1: Always save JSON version to retain full structure editability
        await MSGraphService.saveSpecification(code, formData, newStatus, account.name, originalAuthor);
        
        // Step 2: If HOD approves, automatically generate DOCX and save it alongside
        if (newStatus === 'Approved') {
           showToast("JSON Saved. Generating Official Word Document...", "loading");
           const docxBlob = await DocxEngine.generateBlob(formData);
           
           showToast("Uploading Word Document to OneDrive...", "loading");
           await MSGraphService.uploadDocxBlob(code, docxBlob);
        }

        await refreshDocuments();
        showToast(`Document successfully updated to: ${newStatus}`, "success");
      } else {
        // Local mode fallback
        const newDoc = { 
          id: selectedDocMeta?.id || Date.now(), name: code, courseTitle: formData.basicInfo.courseTitle,
          author: originalAuthor || account.name, lastModified: 'Just now', status: newStatus, data: formData
        };
        setDocuments(prev => [newDoc, ...prev.filter(d => d.id !== newDoc.id)]);
        showToast("Saved locally (Preview Mode).", "success");
      }
      navigate('DASHBOARD');
    } catch (error) {
      showToast(error.message || "Failed to save document. Please try again.", "error");
    }
  };

  const filteredDocuments = useMemo(() => {
    if (!searchTerm) return documents;
    const lower = searchTerm.toLowerCase();
    return documents.filter(d => 
        (d.name || '').toLowerCase().includes(lower) || 
        (d.courseTitle || '').toLowerCase().includes(lower) || 
        (d.author || '').toLowerCase().includes(lower) ||
        (d.status || '').toLowerCase().includes(lower)
    );
  }, [documents, searchTerm]);


if (!isMsalReady) return <div className="min-h-screen flex items-center justify-center text-slate-500 font-sans tracking-wide px-4 text-center">Initializing Identity Services...</div>;

  // STRICT AUTH GUARD: If there is no active account, forcefully render the LOGIN view.
  if (!account || view === 'LOGIN') {
    return (
      <div className="min-h-screen bg-[#f3f2f1] flex items-center justify-center p-4 sm:p-8 font-sans relative overflow-hidden">
        <ToastMessage toast={toast} />
        
        {/* Subtle Watermark BG */}
        <div className="fixed inset-0 flex justify-center items-center pointer-events-none overflow-hidden z-0">
          <img src={logo} alt="Watermark" className="w-[120%] max-w-[400px] sm:w-[80%] sm:max-w-[200px] opacity-[0.05] object-contain" />
        </div>

        <div className="max-w-md w-full relative z-10">
          <div className="text-center mb-8 md:mb-10 flex flex-col items-center">
            <div className="bg-[#1e1e1e] p-3 sm:p-4 px-5 sm:px-6 rounded-2xl shadow-lg mb-5 sm:mb-6 inline-block">
              <img src={msruasLogo} alt="MSRUAS Logo" className="h-16 sm:h-20 w-auto object-contain" />
            </div>
            <h1 className="text-2xl sm:text-3xl font-light text-gray-900 mb-1 sm:mb-2 tracking-tight">Course Portal</h1>
            <p className="text-gray-500 text-xs sm:text-sm">Course Specifications Dashboard</p>
          </div>
          <div className="bg-white p-6 sm:p-8 rounded-xl shadow-xl border-t-4 border-[#0f6cbd] text-center">
            <IconMicrosoft size={48} className="mx-auto mb-4 sm:mb-5" />
            <h2 className="text-lg sm:text-xl font-bold text-gray-800 mb-2">Sign In Required</h2>
            <p className="text-gray-500 text-xs sm:text-sm mb-5 sm:mb-6">Log in with your university account to access the specifications repository.</p>
            <button onClick={handleMicrosoftLogin} className="w-full bg-[#0f6cbd] text-white py-2.5 sm:py-3 rounded-md font-bold hover:bg-[#0c5697] shadow-md transition flex items-center justify-center gap-2 text-sm sm:text-base">
                Connect via Microsoft
            </button>
          </div>
        </div>
      </div>
    );
  }

  const canFacultyEdit = role === 'FACULTY' && (!selectedDocMeta?.status || selectedDocMeta.status === 'Draft' || selectedDocMeta.status === 'Needs Revision');
  const canHodEdit = role === 'HOD';
  const canEdit = canFacultyEdit || canHodEdit;

  if (view === 'VIEWER') {
    return (
      <>
        <ToastMessage toast={toast} />
        <DocumentViewer 
          docMeta={selectedDocMeta} 
          docData={selectedDocData} 
          onBack={() => navigate('DASHBOARD')} 
          onEdit={() => navigate('FORM')} 
          canEdit={canEdit} 
          role={role}
          showToast={showToast} 
          onStatusChange={(data, status) => saveDocument(data, status)}
        />
      </>
    );
  }

  return (
    <div className="min-h-screen bg-[#f8fafc] font-sans relative">
      <ToastMessage toast={toast} />
      
      {/* App Header (Mobile responsive flex) */}
      <header className="bg-[#0f6cbd] text-white px-4 sm:px-6 py-3 flex flex-col sm:flex-row justify-between items-center sticky top-0 z-20 shadow-md gap-3 sm:gap-0">
        <div className="flex items-center gap-3 sm:gap-4 w-full sm:w-auto justify-between sm:justify-start">
          <div className="flex items-center gap-3">
            <div className="bg-white/10 px-2 py-1 rounded-md shadow-sm shrink-0">
              <img src={msruasLogo} alt="MSRUAS Logo" className="h-6 sm:h-8 w-auto object-contain" />
            </div>
            <div>
              <h1 className="text-sm sm:text-lg font-bold tracking-wide">Course Spec Portal</h1>
              <p className="text-[10px] sm:text-xs text-blue-100 font-medium opacity-90">{role === 'HOD' ? 'HOD Workspace' : 'Faculty Workspace'}</p>
            </div>
          </div>
          {/* Logout button moved to top right on mobile */}
          <button onClick={handleLogout} className="sm:hidden bg-blue-800/50 p-2 rounded-full hover:bg-blue-800 transition"><IconLogOut size={16} /></button>
        </div>

        <div className="hidden sm:flex items-center gap-6">
          <div className="text-right">
             <div className="font-bold text-sm">{account?.name || 'User'}</div>
             <div className="text-[10px] text-blue-200 uppercase tracking-wider">Logged in</div>
          </div>
          <button onClick={handleLogout} className="bg-blue-800/50 p-2 rounded-full hover:bg-blue-800 transition"><IconLogOut size={18} /></button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-6 sm:py-10 relative">
        {/* Subtle Dashboard Watermark */}
        <div className="fixed inset-0 flex justify-center items-center pointer-events-none overflow-hidden z-0">
          <img src={logo} alt="Watermark" className="w-[120%] max-w-[400px] sm:w-[80%] sm:max-w-[200px] opacity-[0.05] object-contain" />
        </div>

        {view === 'DASHBOARD' && (
          <div className="space-y-6 sm:space-y-8 relative z-10">
            {/* Dashboard Header & Controls */}
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-gray-200">
              <div>
                <h2 className="text-xl sm:text-2xl font-bold text-gray-900 tracking-tight">Department Repository</h2>
              </div>
              <div className="flex flex-col sm:flex-row gap-3 w-full md:w-auto">
                <div className="relative w-full sm:w-auto">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <IconSearch size={16} className="text-gray-400" />
                  </div>
                  <input 
                    type="text" 
                    placeholder="Search courses..." 
                    className="pl-10 pr-4 py-2 border border-gray-300 rounded-lg w-full sm:w-64 focus:ring-2 focus:ring-[#0f6cbd] focus:border-transparent outline-none text-sm transition"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                </div>
                {role === 'FACULTY' && (
                  <button onClick={openNewForm} className="bg-[#0f6cbd] text-white px-4 sm:px-5 py-2 rounded-lg shadow-md hover:bg-[#0c5697] transition flex items-center justify-center gap-2 font-bold text-sm sm:text-base w-full sm:w-auto whitespace-nowrap">
                    <IconPlus size={18} /> New Spec
                  </button>
                )}
              </div>
            </div>

            {/* Document Grid */}
            {toast?.type === 'loading' && documents.length === 0 ? null : filteredDocuments.length === 0 ? (
              <div className="bg-white/80 backdrop-blur-sm p-8 sm:p-16 text-center rounded-xl border border-gray-200 shadow-sm">
                <div className="bg-gray-50 w-16 h-16 sm:w-20 sm:h-20 rounded-full flex items-center justify-center mx-auto mb-4 border border-gray-200"><IconCloud size={28} className="text-gray-400" /></div>
                <h3 className="text-base sm:text-lg font-bold text-gray-800">No documents found</h3>
                <p className="text-gray-500 mt-1 text-xs sm:text-sm">{searchTerm ? "Try adjusting your search filters." : "Your department repository is currently empty."}</p>
                {(!searchTerm && role === 'FACULTY') && <button onClick={openNewForm} className="mt-6 text-[#0f6cbd] text-sm sm:text-base font-bold hover:underline">Create your first spec &rarr;</button>}
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-6">
                {filteredDocuments.map(doc => {
                   const rowCanFacultyEdit = role === 'FACULTY' && (doc.status === 'Draft' || doc.status === 'Needs Revision');
                   const rowCanEdit = role === 'HOD' || rowCanFacultyEdit;

                   return (
                     <div key={doc.id} className="bg-white/95 backdrop-blur-sm rounded-xl shadow-sm hover:shadow-md transition-shadow border border-gray-200 overflow-hidden flex flex-col group relative">
                       {/* Status Banner */}
                       <div className={`h-1.5 w-full ${
                          doc.status === 'Approved' ? 'bg-green-500' : 
                          doc.status === 'Needs Revision' ? 'bg-yellow-400' : 
                          doc.status === 'Draft' ? 'bg-gray-300' : 'bg-[#0f6cbd]'
                       }`}></div>
                       
                       <div className="p-4 sm:p-6 flex-1">
                         <div className="flex justify-between items-start mb-3">
                           <span className={`px-2 py-1 text-[9px] sm:text-[10px] font-bold uppercase tracking-wider rounded-md border ${
                              doc.status === 'Approved' ? 'bg-green-50 text-green-700 border-green-200' : 
                              doc.status === 'Needs Revision' ? 'bg-yellow-50 text-yellow-700 border-yellow-200' : 
                              doc.status === 'Draft' ? 'bg-gray-100 text-gray-600 border-gray-200' : 
                              'bg-blue-50 text-blue-700 border-blue-200'
                           }`}>
                             {doc.status}
                           </span>
                         </div>
                         
                         <h3 className="font-bold text-gray-900 text-base sm:text-lg mb-1 truncate" title={doc.courseTitle}>{doc.name}</h3>
                         <p className="text-xs sm:text-sm text-gray-500 font-medium truncate mb-4">{doc.courseTitle}</p>
                         
                         <div className="space-y-1.5 sm:space-y-2 text-[10px] sm:text-xs text-gray-500 mt-auto pt-3 sm:pt-4 border-t border-gray-100">
                           <div className="flex items-center justify-between">
                             <span>Author:</span>
                             <span className="font-medium text-gray-900 truncate max-w-[100px] sm:max-w-[120px]">{doc.author}</span>
                           </div>
                           <div className="flex items-center justify-between">
                             <span>Modified:</span>
                             <span className="font-medium text-gray-900">{doc.lastModified}</span>
                           </div>
                         </div>
                       </div>
                       
                       <div className="bg-gray-50/80 border-t border-gray-100 p-3 flex justify-end gap-2">
                         <button onClick={() => openDocumentViewer(doc)} className="flex-1 sm:flex-none px-3 py-1.5 text-[#0f6cbd] hover:bg-blue-50 font-bold text-xs rounded transition flex justify-center items-center gap-1.5">
                           <IconEye size={14}/> View
                         </button>
                         {rowCanEdit && (
                           <button onClick={() => {
                             setSelectedDocMeta(doc);
                             setSelectedDocData(doc.data || emptyFormState); 
                             openDocumentViewer(doc).then(() => navigate('FORM'));
                           }} className="flex-1 sm:flex-none px-3 py-1.5 text-gray-700 hover:bg-gray-200 font-bold text-xs rounded transition flex justify-center items-center gap-1.5 border border-gray-300 bg-white shadow-sm">
                             <IconEdit size={14}/> Edit
                           </button>
                         )}
                       </div>
                     </div>
                   );
                })}
              </div>
            )}
          </div>
        )}

        {view === 'FORM' && (
          <div className="relative z-10">
            <CourseForm 
               initialData={selectedDocData} 
               currentStatus={selectedDocMeta?.status || 'Draft'}
               role={role}
               onSave={saveDocument} 
               onCancel={() => navigate('DASHBOARD')} 
            />
          </div>
        )}
      </main>
    </div>
  );
}