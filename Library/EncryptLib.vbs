Option Explicit

''' #########################################################
''' <summary>
''' Demonstrate a simple encryption algorithm.
''' </summary>
''' <remarks></remarks>	 
''' <example>
''' Dim strEncrypt, strDecrypt, strEncryptedText
''' Tis is the text that is going to be encrypted
''' strEncrypt = "1375819679"
''' Call the Encrypt function to encrypt the text
''' strEncryptedText = OEncryptLib.Encrypt(strEncrypt)
''' MsgBox strEncryptedText
''' strDecrypt = strEncryptedText
''' Output the decrypted text to screen
''' MsgBox "Decrypted text : " & OEncryptLib.Decrypt(strDecrypt) 
''' </example>
''' #########################################################

Class ClsEncryptLib
	
	Private strKey
	Private intSeed
	
	''' <summary>
    ''' Class Initialization procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
			
		'This is the key that will be use for en/decrypting the text
		strKey = "ABCDEFGHIJKLMN123456789abcdefghigklmn"
		'This is the seed that is used for randpmizing the en/decryption
		intSeed = 1
		
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()

	End Sub
	
	''' <summary>
	''' Transform a string in to an array with ascii values
	''' </summary>
	''' <param name="strIn" type="string">string to be converted
	''' <returns>an array with ascii values</returns>    
	''' <remarks></remarks>
	Public Function String2Asc( strIn)
		
		Dim arrResult, intI
		arrResult = Array()
		ReDim arrResult( CInt( Len( strIn ) ) )
		For intI = 0 to Len(strIn) - 1
		  arrResult( intI ) = Asc( Mid( strIn,intI + 1 ,1 ) )
		Next
		String2Asc = arrResult
		
	End Function

	''' <summary>
	''' Encrypt a string in to an encrypted string
	''' </summary>
	''' <param name="strEncrypt" type="string">string to be encrypted
	''' <returns>an encrypted string</returns>    
	''' <remarks></remarks>
	Public Function Encrypt(ByVal strEncrypt)
	  	
	  	Dim intRnd, intI, intPointer,intCalc, arrEncrypt, arrKey, strEncrypted
		Rnd(-1)
		Randomize intSeed
		intRnd =  Int( ( Len(strKey) - 1 + 1 ) * Rnd + 1 )
		
		arrEncrypt = String2Asc(strEncrypt)
		arrKey = String2Asc(strKey)
		
		For intI = 0 to UBound( arrEncrypt ) - 1
		  
		  intPointer = intI + intRnd
		  If intPointer > UBound(arrKey) Then
		     intPointer = intPointer -  ((UBound(arrKey) + 1 ) * Int(intPointer / (UBound(arrKey) + 1)))
		  End If
		  
		  intCalc = arrEncrypt(intI) + arrKey(intPointer)
		  
		  If intCalc > 256 Then
		  	intCalc = intCalc - 256 
		  End If
		  strEncrypted = strEncrypted & Chr(intCalc)
		Next
		encrypt = strEncrypted
	  
	End Function
    
	''' <summary>
	''' Decrypt an encrypted string
	''' </summary>
	''' <param name="strDecrypt" type="string">string to be Decrypted
	''' <returns>A Decrypted string</returns>    
	''' <remarks></remarks>
	Function Decrypt(ByVal strDecrypt)
	  
	  	Dim intRnd, intI, intPointer, intCalc, arrDecrypt, arrKey, strDecrypted
		Rnd(-1)
		Randomize intSeed
		intRnd =  Int( ( Len(strKey) - 1 + 1 ) * Rnd + 1 )
		
		arrDecrypt = String2Asc(strDecrypt)
		arrKey = String2Asc(strKey)
		
		For intI = 0 to UBound( arrDecrypt ) - 1
		  
		  intPointer = intI + intRnd
		  If intPointer > UBound(arrKey) Then
		     intPointer = intPointer -  ((UBound(arrKey) + 1 ) * Int(intPointer / (UBound(arrKey) + 1)))
		  End If
		  
		  intCalc = arrDecrypt(intI) - arrKey(intPointer)
		  
		  If intCalc < 0 Then
		  	intCalc = intCalc + 256 
		  End If
		  strDecrypted = strDecrypted & Chr(intCalc)
		Next
		Decrypt = strDecrypted
		
	End Function
	
End Class

Public Function EncryptLib()
	
	Dim objEncryptLib
	Set objEncryptLib = New ClsEncryptLib
	Set EncryptLib = objEncryptLib

End Function