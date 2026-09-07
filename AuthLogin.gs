// ==========================================
// 🧠 CENTRAL AUTH + RBAC ENGINE
// ==========================================

function verifyLogin(credentials) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usernameInput = credentials.username.toString().trim();
  const passwordInput = credentials.password.toString().trim();

  // ==========================================
  // 1. HOSPITAL STAFF LOGIN
  // ==========================================
  const userSheet = ss.getSheetByName('Users');

  if (userSheet) {
    const userData = userSheet.getDataRange().getValues();

    for (let i = 1; i < userData.length; i++) {
      const storedUsername = userData[i][0];
      const storedPassword = userData[i][1];
      const storedRole = userData[i][2];
      const isActive = userData[i][3];

      if (!storedUsername) continue;

      if (storedUsername.toString().trim().toUpperCase() === usernameInput.toUpperCase()) {
        
        // Check Active Status
        if (isActive && isActive.toString().toLowerCase() !== 'active') {
          return { success: false, message: "Account disabled." };
        }

        // Password Validation
        if (storedPassword.toString().trim() === passwordInput) {
          const role = storedRole.toString().trim().toLowerCase();
          const doc  = resolveDoctorByUsername_(storedUsername);
          const token = issueSession_({
            username: storedUsername,
            role: role,
            doctorId: doc ? doc.doctorId : "",
            name: doc ? doc.name : storedUsername
          });
          return {
            success: true,
            role: role,
            portal: 'hospital',
            username: storedUsername,
            displayName: doc ? doc.name : storedUsername,
            doctorId: doc ? doc.doctorId : "",
            sessionToken: token,
            message: "Welcome " + storedRole
          };
        } else {
          return { success: false, message: "Incorrect staff password." };
        }
      }
    }
  }

  // ==========================================
  // 2. PATIENT LOGIN (DYNAMIC PASSWORD ENGINE)
  // ==========================================
  const patientSheet = ss.getSheetByName('Patients');
  if (!patientSheet) {
    return { success: false, message: "Patients database missing." };
  }

  const patientData = patientSheet.getDataRange().getValues();

  for (let i = 1; i < patientData.length; i++) {
    const patientID = patientData[i][0];
    if (!patientID) continue;

    if (patientID.toString().trim().toUpperCase() === usernameInput.toUpperCase()) {
      
      let rawName = patientData[i][2] ? patientData[i][2].toString().trim() : "XXX";
      let rawDOB = patientData[i][5];

      // Format Name Part (First 3 chars)
      let namePart = rawName.replace(/[^a-zA-Z]/g, '');
      if (namePart.length < 3) {
        namePart = (namePart + "XXX").substring(0, 3);
      } else {
        namePart = namePart.substring(0, 3);
      }
      namePart = namePart.charAt(0).toUpperCase() + namePart.substring(1).toLowerCase();

      // Format Year Part (4 Digits)
      let yearPart = "0000";
      if (rawDOB instanceof Date) {
        yearPart = rawDOB.getFullYear().toString();
      } else if (rawDOB) {
        let dobStr = rawDOB.toString().trim();
        let yearMatch = dobStr.match(/\b(19|20)\d{2}\b/);
        if (yearMatch) {
          yearPart = yearMatch[0];
        } else {
          yearPart = dobStr.length >= 4 ? dobStr.slice(-4) : "0000";
        }
      }

      const expectedPassword = namePart + yearPart;

      // Validate Password
      if (passwordInput.trim().toLowerCase() === expectedPassword.toLowerCase()) {
        return {
          success: true,
          role: 'patient',
          portal: 'patient',
          message: "Welcome Patient"
        };
      } else {
        return { success: false, message: "Incorrect password. Format: Name(3 chars) + Birth Year." };
      }
    }
  }

  // ==========================================
  // 3. USER NOT FOUND
  // ==========================================
  return { success: false, message: "User ID not found." };
}

// ==========================================
// GOOGLE SSO AUTHENTICATION ENGINE
// ==========================================
function verifyGoogleLogin(userEmail) {
  try {
    if (!userEmail) {
      return { success: false, message: "Authentication payload missing email link." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) throw new Error("Users staff registry tab is missing.");

    const userData = userSheet.getDataRange().getValues();

    // Check email alignment against column E (index 4) of your Users dataset
    for (let i = 1; i < userData.length; i++) {
      const storedUsername = userData[i][0];
      const storedRole = userData[i][2];
      const isActive = userData[i][3];
      const storedEmail = userData[i][4];

      if (!storedEmail) continue;

      if (storedEmail.toString().trim().toLowerCase() === userEmail.toString().trim().toLowerCase()) {
        
        if (isActive && isActive.toString().toLowerCase() !== 'active') {
          return { success: false, message: "Access Denied: This staff credential context is deactivated." };
        }

        const role = storedRole.toString().trim().toLowerCase();
        const doc  = resolveDoctorByUsername_(storedUsername);
        const token = issueSession_({
          username: storedUsername,
          role: role,
          doctorId: doc ? doc.doctorId : "",
          name: doc ? doc.name : storedUsername
        });
        return {
          success: true,
          role: role,
          portal: 'hospital',
          username: storedUsername,
          displayName: doc ? doc.name : storedUsername,
          doctorId: doc ? doc.doctorId : "",
          sessionToken: token,
          message: "Welcome back " + storedUsername
        };
      }
    }

    return { 
      success: false, 
      message: "Access Denied: The email " + userEmail + " is not registered in CresRx. Contact Admin." 
    };

  } catch (error) {
    return { success: false, message: "System Security Fault: " + error.toString() };
  }
}

// ==========================================
// MFA / TOTP ENGINE (Google Authenticator)
// ==========================================

function verifyMFA(username, userCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0].toString().trim().toUpperCase() === username.toUpperCase()) {
      const storedSecret = userData[i][5]; // Column F where MFA_Secret is stored
      
      if (!storedSecret) return { success: true }; // Skip MFA if not enrolled
      
      return processTOTP(storedSecret, userCode);
    }
  }
  return { success: false, message: "User not found for MFA verification." };
}

function processTOTP(secretBase32, userToken) {
  try {
    var keyBytes = base32ToBytes(secretBase32);
    var epoch = Math.floor(Date.now() / 1000);
    var timeWindow = Math.floor(epoch / 30);
    
    // Check current, previous, and next 30-second windows (allows for slight clock drift)
    for (var i = -1; i <= 1; i++) {
      if (generateTOTPAlgorithm(keyBytes, timeWindow + i) === String(userToken).trim()) {
        return { success: true, message: "MFA Verified" };
      }
    }
    return { success: false, message: "Invalid or expired 6-digit code." };
  } catch (e) {
    return { success: false, message: "Verification error." };
  }
}

function generateTOTPAlgorithm(keyBytes, timeValue) {
  var timeBytes = new Array(8);
  for (var i = 7; i >= 0; i--) {
    timeBytes[i] = timeValue & 0xff;
    timeValue >>= 8;
  }
  var hmac = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, timeBytes, keyBytes);
  var offset = hmac[hmac.length - 1] & 0x0f;
  var binary = ((hmac[offset] & 0x7f) << 24) | ((hmac[offset + 1] & 0xff) << 16) | ((hmac[offset + 2] & 0xff) << 8) | (hmac[offset + 3] & 0xff);
  var otp = (binary % 1000000).toString();
  while (otp.length < 6) { otp = '0' + otp; }
  return otp;
}

function base32ToBytes(base32) {
  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  var bits = 0, value = 0, index = 0, output = [];
  for (var i = 0; i < base32.length; i++) {
    var val = alphabet.indexOf(base32.charAt(i).toUpperCase());
    if (val === -1) continue; 
    value = (value << 5) | val;
    bits += 5;
    if (bits >= 8) {
      output.push((value >>> (bits - 8)) & 255);
      bits -= 8;
    }
  }
  return output;
}