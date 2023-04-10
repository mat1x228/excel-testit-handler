document.getElementById("processBtn").addEventListener("click", function() { 
  const file = document.getElementById("fileInput").files[0]; 
  const reader = new FileReader(); 

  reader.onload = function(e) { 
      const data = e.target.result; 
      const workbook = XLSX.read(data, { type: "binary" }); 
      const sheetName = workbook.SheetNames[0]; 
      const worksheet = workbook.Sheets[sheetName]; 
      const json = XLSX.utils.sheet_to_json(worksheet); 

      const ids = []; 

      json.forEach((row) => { 
          const id = row.ID; 
          if (id !== undefined) { 
              ids.push(id); 
          } 
      }); 

      window.ids = ids; 
      console.log(window.ids); 
  }; 

  reader.readAsBinaryString(file); 
}); 

document.getElementById("generateBtn").addEventListener("click", function() { 
  const projectId = document.getElementById("projectId").value;
  const ws_name = "Sheet1"; 
  const wb = XLSX.utils.book_new(); 
  const ws_data = [ 
      ["Use cases", "Критичность", ...window.ids.map((id) => ({ t: "s", v: id, l: { Target: `https://testit.com/projects/${projectId}/tests/${id}` } }))], 
  ]; 
  const ws = XLSX.utils.aoa_to_sheet(ws_data); 
  XLSX.utils.book_append_sheet(wb, ws, ws_name); 

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" }); 
  saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), "output.xlsx"); 
}); 


function s2ab(s) { 
  const buf = new ArrayBuffer(s.length); 
  const view = new Uint8Array(buf); 
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff; 
  return buf; 
}
