#tag Class
Protected Class NVirtualMachineDetector
	#tag Method, Flags = &h21
		Private Sub BuildInfo()
		  if mInfos.Ubound > -1 then return
		  
		  mInfos.Append "Virtualized: " + Str( mIsVirtualized )
		  mInfos.Append "Product: " + if ( mVMProductName = "", "n/a", mVMProductName )
		  mInfos.Append "Vendor: " + if ( mVMVendor = "", "n/a", mVMVendor )
		  mInfos.Append "BIOS Vendor: " + if ( mBIOSVendor = "", "n/a", mBIOSVendor )
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  #if TargetMacOS
		    DetectVirtualizedMacOS()
		    
		  #elseif TargetWindows
		    DetectVirtualizedWindows()
		    
		  #else
		    mInfos.Append "Unsupported OS"
		    
		  #endif
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub DetectVirtualizedMacOS()
		  #if TargetMacOS
		    
		    mInfos.Append "OS X/macOS not supported yet"
		    BuildInfo()
		    
		  #endif
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Attributes( "in work" ) Private Sub DetectVirtualizedWindows()
		  #if TargetWindows
		    
		    if xC___WIN_ParallelsVirtualPlatform() then goto detected
		    if xC___WIN_VMware() then goto detected
		    if xC___WIN_VirtualBox() then goto detected
		    if xC___WIN_Sandboxed() then goto detected
		    
		    detected:
		    BuildInfo()
		    
		  #endif
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Info() As string
		  return Join( mInfos, EndOfLine )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function InfoStr() As string
		  dim infos() as string
		  
		  if mIsVirtualized or mIsSandboxed then
		    if mVMProductName <> "" then infos.Append mVMProductName
		    if mVMVendor <> "" then infos.Append mVMVendor
		    if mBIOSVendor <> "" then infos.Append mBIOSVendor
		    
		    if infos.Ubound = -1 then
		      if mIsVirtualized then
		        infos.Append "Virtualized"
		        
		      elseif mIsSandboxed then
		        infos.Append "Sandboxed"
		        
		      else
		        break
		        
		      end if
		      
		    end if
		  end if
		  
		  return Join( infos, ", " ).Trim()
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MachineIsSandboxed() As boolean
		  return mIsSandboxed
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MachineIsVirtualized() As boolean
		  return mIsVirtualized
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function WinRegistrySearch(reg_key_ as RegistryItem, property_ as string, value_ as string, level_ as integer) As boolean
		  #if TargetMacOS
		    #pragma unused reg_key_
		    #pragma unused property_
		    #pragma unused value_
		    #pragma unused level_
		  #endif
		  
		  #if TargetWindows
		    
		    dim names() as string
		    dim values() as string
		    dim index as integer
		    
		    for i as integer = 0 to reg_key_.KeyCount - 1
		      
		      names.Append reg_key_.Name( i )
		      values.Append reg_key_.Value( i )
		      
		      if names( names.Ubound ).Left( property_.Len() ) = property_ then
		        if value_ = "" then return true
		        
		        if values( values.Ubound ).Left( value_.Len() ) = value_ then
		          return true
		          
		        end if
		      end if
		      
		    next
		    
		    
		    
		    dim found as boolean
		    dim next_reg_key as RegistryItem
		    
		    for i as integer = 0 to reg_key_.FolderCount - 1
		      
		      #pragma BreakOnExceptions off
		      try
		        next_reg_key = reg_key_.Item( i )
		        if next_reg_key <> nil then
		          found = WinRegistrySearch( next_reg_key, property_, value_, level_ + 1 )
		          
		          if found then return true
		          
		        end if
		        
		      catch ex as RuntimeException
		        //
		        // ignore
		        //
		        break
		        
		      end try
		      #pragma BreakOnExceptions default
		      
		    next
		    
		  #endif
		  
		  return false
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Private Function xC___WIN_ParallelsVirtualPlatform() As boolean
		  #if TargetWindows
		    //
		    // check for Paralles Virtual Platform
		    //
		    static REC_STR as string = "parallels"
		    
		    try
		      dim registry_item as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS" )
		      dim names() as string
		      dim values() as string
		      dim index as integer
		      
		      for i as integer = 0 to registry_item.KeyCount - 1
		        
		        names.Append registry_item.Name( i )
		        values.Append registry_item.Value( i )
		        
		      next
		      
		      index = names.IndexOf( "BIOSVendor" )
		      if index > -1 then mBIOSVendor = values( index )
		      
		      index = names.IndexOf( "SystemManufacturer" )
		      if index > -1 then
		        if values( index ).Left( REC_STR.Len() ).Lowercase() = REC_STR then
		          mIsVirtualized = true
		          mVMVendor = values( index )
		          
		          index = names.IndexOf( "SystemProductName" )
		          if index > -1 then mVMProductName = values( index )
		          
		        end if
		      end if
		      
		    catch
		      //
		      // ignore
		      //
		    end try
		    
		    
		  #endif
		  
		  return mIsVirtualized
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Attributes( "in work" ) Private Function xC___WIN_Sandboxed() As boolean
		  //@mod 2016-09-08, cb 
		  
		  #if TargetWindows
		    
		    static SB_ANUBIS as string = "76487-337-8429955-22614"
		    static SB_CWSANDBOX as string = "76487-644-3177037-23510"
		    static SB_JOEBOX as string = "55274-640-2673064-23950"
		    
		    try
		      dim registry_item as new RegistryItem( "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion" )
		      
		      if WinRegistrySearch( registry_item, "PackageMoniker", "Syste", 0 ) then
		        mIsSandboxed = true
		        mVMProductName = "!"
		        mVMVendor = "?"
		        
		        return true
		        
		      end if
		      
		      if WinRegistrySearch( registry_item, "ProductId", SB_ANUBIS, 0 ) then
		        mIsSandboxed = true
		        mVMProductName = "Anubis"
		        mVMVendor = "International Secure Systems Lab"
		        
		        return true
		        
		      end if
		      
		      if WinRegistrySearch( registry_item, "ProductId", SB_CWSANDBOX, 0 ) then
		        mIsSandboxed = true
		        mVMProductName = "CWSandbox"
		        mVMVendor = "Sunbelt Software"
		        
		      end if
		      
		      if WinRegistrySearch( registry_item, "ProductId", SB_JOEBOX, 0 ) then
		        mIsSandboxed = true
		        mVMProductName = "Joe Sandbox"
		        mVMVendor = "Joe Security LLC"
		        
		      end if
		      
		    catch
		      //
		      // ignore
		      //
		    end try
		    
		  #endif
		  
		  return mIsSandboxed
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Private Function xC___WIN_VirtualBox() As boolean
		  #if TargetWindows
		    //
		    // check for Oracle VirtualBox
		    //
		    static REC_STR as string = "vbox"
		    
		    try
		      dim registry_item as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System" )
		      dim names() as string
		      dim values() as string
		      dim index as integer
		      
		      for i as integer = 0 to registry_item.KeyCount - 1
		        
		        names.Append registry_item.Name( i )
		        values.Append registry_item.Value( i )
		        
		      next
		      
		      index = names.IndexOf( "SystemBiosVersion" )
		      if index > -1 then
		        if values( index ).Left( REC_STR.Len() ).Lowercase() = REC_STR then
		          mIsVirtualized = true
		          mVMProductName = "VirtualBox"
		          mVMVendor = "Oracle Corporation"
		          
		          //
		          index = names.IndexOf( "VideoBiosVersion" )
		          if index > -1 then mBIOSVendor = values( index )
		          
		        end if
		      end if
		      
		    catch
		      //
		      // ignore
		      //
		    end try
		    
		    
		  #endif
		  
		  return mIsVirtualized
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = (TargetConsole and (Target32Bit or Target64Bit)) or  (TargetWeb and (Target32Bit or Target64Bit)) or  (TargetDesktop and (Target32Bit or Target64Bit))
		Private Function xC___WIN_VMware() As boolean
		  #if TargetWindows
		    //
		    // check for VMware
		    //
		    static REC_STR as string = "vmware"
		    
		    try
		      dim registry_item as new RegistryItem( "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS" )
		      dim names() as string
		      dim values() as string
		      dim index as integer
		      
		      for i as integer = 0 to registry_item.KeyCount - 1
		        
		        names.Append registry_item.Name( i )
		        values.Append registry_item.Value( i )
		        
		      next
		      
		      index = names.IndexOf( "BIOSVendor" )
		      if index > -1 then mBIOSVendor = values( index )
		      
		      index = names.IndexOf( "SystemManufacturer" )
		      if index > -1 then
		        if values( index ).Left( REC_STR.Len() ).Lowercase() = REC_STR then
		          mIsVirtualized = true
		          mVMVendor = values( index )
		          
		          index = names.IndexOf( "SystemProductName" )
		          if index > -1 then mVMProductName = values( index )
		          
		        end if
		      end if
		      
		    catch
		      //
		      // ignore
		      //
		    end try
		    
		    
		  #endif
		  
		  return mIsVirtualized
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private mBIOSVendor As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mInfos() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsSandboxed As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mIsVirtualized As boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVMProductName As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mVMVendor As string
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
