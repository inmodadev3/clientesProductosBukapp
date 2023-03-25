const axios = require('axios')
const xl = require('excel4node')
const path = require('path')
const fs = require('fs')

//Fechas 
const fecha = new Date()
const ano = fecha.getFullYear()
const mes = fecha.getMonth()
const dia = fecha.getDate()
const hora = fecha.getHours()
const minuto = fecha.getMinutes()
let segundo  = fecha.getSeconds()
if(segundo < 10){
    segundo = `0${segundo}`
}
const docFecha = `${ano}-0${mes + 1}-${dia}-${hora}-${minuto}-${segundo}`

const consultarProductos = () => {
    //excel
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('ProductosActivos');
    let style = wb.createStyle({
        font: {
            color: '#040404',
            size: 12
        }
    })
    let red = wb.createStyle({
        font: {
            color: '#FF0000',
            size: 18
        }
    })

    ws.cell(1, 1).string("StrIdProductos").style(red)

    axios.post('https://api.bukappweb.com:3000/api/usuarios/login', {
        strIdUsuario: 'sistemas2inmoda@gmail.com',
        strClave: 'Inmoda2021*'
    }).then((response) => {
        if (response.data.success) {
            let user = response.data.data
            axios.get('https://api.bukappweb.com:3000/api/productos/images/', {
                headers: {
                    authorization: `Bearer ${user.token}`,
                    auth: JSON.stringify({
                        usuario_id: user.id,
                        empresa_id: user.empresa_id,
                        where: {
                            ...(null ? {} : { boolHabilitado: true })
                        },
                    })
                }
            }).then((response) => {
                let productos = response.data.data
                if (response.data.success) {
                    productos.forEach((producto, index) => {
                        ws.cell(index + 2, 1).string(producto.strIdProducto).style(style)
                        ws.cell(index + 2, 2).string(`'${producto.strIdProducto}',`).style(style)
                    });

                    const pathExcel = path.join(__dirname, 'excel/productos', `ProductosHabilitados-${docFecha}.xlsx`);

                    wb.write(pathExcel, (err, stats) => {
                        if (err) {
                            console.log(err)
                        } else {
                            console.log('SE HA GENERADO EL EXCEL DE PRODUCTOS CORRECTAMENTE')
                        }
                    })
                }
            }).catch((err) => {
                console.log(err)
            })
        }
    }).catch((err) => {
        console.log(err)
    })
}


const Consultar_clientes = () => {
    let wb = new xl.Workbook();
    let ws = wb.addWorksheet('Clientes');
    let dataCliente = wb.addWorksheet('dataCliente')
    let style = wb.createStyle({
        font: {
            color: '#040404',
            size: 12
        }
    })
    let red = wb.createStyle({
        font: {
            color: '#FF0000',
            size: 11
        }
    })

    ws.cell(1, 1).string("strIdTercero").style(red)
    ws.cell(1, 2).string("strNombre").style(red)
    ws.cell(1, 3).string("strTelefono").style(red)
    ws.cell(1, 4).string("strCelularTercero").style(red)
    ws.cell(1, 5).string("strEmailTercero").style(red)
    ws.cell(1, 6).string("strDireccionTercero").style(red)

    dataCliente.cell(1,1).string("strIdTercero").style(red)
    dataCliente.cell(1,2).string("strNombre").style(red)
    dataCliente.cell(1,3).string("strPassword").style(red)



    axios.post('https://api.bukappweb.com:3000/api/usuarios/login', {
        strIdUsuario: 'sistemas2inmoda@gmail.com',
        strClave: 'Inmoda2021*'
    }).then((response) => {
        const user = response.data.data;
        console.log("Iniciando...")
        axios.get('https://api.bukappweb.com:3000/api/terceros/', {
            headers: {
                authorization: `Bearer ${user.token}`,
                auth: JSON.stringify({
                    usuario_id: user.id,
                    empresa_id: user.empresa_id,
                })
            }
        }).then((response) => {
            let Clientes = response.data.data
            if(response.data.success){
                Clientes.forEach((Cliente,index) => {
                    ws.cell(index+2,1).string(Cliente.strIdTercero).style(style)
                    ws.cell(index+2,2).string(`'${Cliente.strNombre}',`).style(style)
                    ws.cell(index+2,3).string(`'${Cliente.strTelefono}',`).style(style)
                    ws.cell(index+2,4).string(`'${Cliente.strCelular}',`).style(style)
                    ws.cell(index+2,5).string(`'${Cliente.strEmail}',`).style(style)
                    ws.cell(index+2,6).string(`'${Cliente.strDireccion}',`).style(style)

                    dataCliente.cell(index+2,1).string(Cliente.strIdTercero).style(style)
                    dataCliente.cell(index+2,2).string(Cliente.strNombre).style(style)
                    dataCliente.cell(index+2,3).string(Cliente.strPassword).style(style)
                });

                const pathExcel = path.join(__dirname,'excel/clientes',`ClientesLista-${docFecha}.xlsx`);

                wb.write(pathExcel,(err,stats)=>{
                    if(err){
                        console.log(err)
                    }else{
                        console.log('SE HA GENERADO EL EXCEL DE CLIENTES CORRECTAMENTE')
                        setTimeout(() => {
                            
                        }, 4000);
                    }
                })
            }

        }).catch((err) => {
            /* console.log(err) */
            console.log('error 2')
        })
    }).catch((err) => {
        console.log('error 1')
        /* console.log(err) */
    })

}

consultarProductos()
Consultar_clientes()
