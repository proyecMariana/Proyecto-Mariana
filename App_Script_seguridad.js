function searchUser(){
	//Obteniendo la sesion activa del correo institucional del usuario de la unviersidad Mariana
	const activeUser = Session.getActiveUser().getEmail();
	//Metodo que nos permite conectarnos a una hoja de un archivo de excel
	const SS = SpreadsheetApp.getActiveSpreadsheet();
	//Usuarios corresponde al nombre de la hoja de excel
	const sheetUsers = SS.getSheetByName('Usuarios');
	const activeUserList = sheetUsers.getRange(2,1, sheetUsers.getLastRow()-1,1).getValues().map(user => user[0]);
	//console.log(activeUserList+'-->'+activeUser);
	//Permite realizar la busquedad del correo activo en la lista de usuarios habilitados para acceder al aplicativo
	if(activeUserList.indexOf(activeUser)!==-1){
	 //console.log('Dar acceso');
	 return true;
	}else{
    //console.log('No dar acceso');
		return false;
	}
}