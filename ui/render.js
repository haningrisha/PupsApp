let path  = document.getElementById("input-folder").files[0].path;
const request = net.request({
  method: 'POST',
  protocol: 'https:',
  hostname: 'localhost',
  port: 5050,
  path: '/folder'
})

