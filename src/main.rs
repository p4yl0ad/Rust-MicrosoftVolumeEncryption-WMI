use windows::{
    core::*, Win32::System::Com::*, Win32::System::Ole::*, Win32::System::Variant::*,
    Win32::System::Wmi::*,
};

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(
            None,
            COINIT_MULTITHREADED
        )?;

        CoInitializeSecurity(
            None,
            -1,
            None,
            None,
            RPC_C_AUTHN_LEVEL_DEFAULT,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            None,
            EOAC_NONE,
            None,
        )?;

        let locator: IWbemLocator = CoCreateInstance(
            &WbemLocator,
            None,
            CLSCTX_INPROC_SERVER
        )?;


        /*
            Try and query Win32_LogicalDisk from cimv2
        */
        let server =
            locator.ConnectServer(
                &BSTR::from("root\\cimv2"),
                None,
                None,
                None,
                0,
                None,
                None
            )?;

        let query = server.ExecQuery(
            &BSTR::from("WQL"),
            &BSTR::from("select Caption from Win32_LogicalDisk"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            None,
        )?;

        loop {
            let mut row = [None; 1];
            let mut returned = 0;
            query.Next(
                WBEM_INFINITE,
                &mut row,
                &mut returned
            ).ok()?;

            if let Some(row) = &row[0] {

                let mut value:VARIANT = Default::default();

                row.Get(
                    w!("Caption"),
                    0,
                    &mut value,
                    None,
                    None
                )?;
                println!(
                    "[i] Drive Letter\t: {}",
                    VarFormat(
                        &value,
                        None,
                        VARFORMAT_FIRST_DAY_SYSTEMDEFAULT,
                        VARFORMAT_FIRST_WEEK_SYSTEMDEFAULT,
                        0
                    )?
                );


                /*
                    Try and query Win32_EncryptableVolume from MicrosoftVolumeEncryption
                */

                let fuck =
                    format!("SELECT * FROM Win32_EncryptableVolume Where DriveLetter='{}'", VarFormat(
                        &value,
                        None,
                        VARFORMAT_FIRST_DAY_SYSTEMDEFAULT,
                        VARFORMAT_FIRST_WEEK_SYSTEMDEFAULT,
                        0
                    )?.to_string());

                println!("[*] WMI Query\t\t: {}", fuck);

                let server2 =
                    locator.ConnectServer(
                        &BSTR::from("root\\cimv2\\Security\\MicrosoftVolumeEncryption"),
                        None,
                        None,
                        None,
                        0,
                        None,
                        None
                    )?;

                let query2 = server2.ExecQuery(
                    &BSTR::from("WQL"),
                    &BSTR::from(fuck.as_str()),
                    WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                    None,
                )?;

                loop {
                    let mut row2 = [None; 1];
                    let mut returned2 = 0;
                    query2.Next(
                        WBEM_INFINITE,
                        &mut row2,
                        &mut returned2
                    ).ok()?;

                    if let Some(row2) = &row2[0] {
                        let mut value2 = Default::default();

                        row2.Get(
                            w!("DeviceID"),
                            0,
                            &mut value2,
                            None,
                            None
                        )?;
                        println!(
                            "[*] DeviceID\t\t: {}",
                            VarFormat(
                                &value2,
                                None,
                                VARFORMAT_FIRST_DAY_SYSTEMDEFAULT,
                                VARFORMAT_FIRST_WEEK_SYSTEMDEFAULT,
                                0
                            )?
                        );

                        row2.Get(
                            w!("PersistentVolumeID"),
                            0,
                            &mut value2,
                            None,
                            None
                        )?;
                        println!(
                            "[*] PersistentVolumeID\t: {}",
                            VarFormat(
                                &value2,
                                None,
                                VARFORMAT_FIRST_DAY_SYSTEMDEFAULT,
                                VARFORMAT_FIRST_WEEK_SYSTEMDEFAULT,
                                0
                            )?
                        );

                        row2.Get(
                            w!("DriveLetter"),
                                 0,
                                 &mut value2,
                                 None,
                                 None
                        )?;
                        println!(
                            "[*] DriveLetter\t\t: {}",
                            VarFormat(
                                &value2,
                                None,
                                VARFORMAT_FIRST_DAY_SYSTEMDEFAULT,
                                VARFORMAT_FIRST_WEEK_SYSTEMDEFAULT,
                                0
                            )?
                        );

                        row2.Get(
                            w!("ProtectionStatus"),
                                 0,
                                 &mut value2,
                                 None,
                                 None
                        )?;
                        println!(
                            "[*] ProtectionStatus\t: {}",
                            VarFormat(
                                &value2,
                                None,
                                VARFORMAT_FIRST_DAY_SYSTEMDEFAULT,
                                VARFORMAT_FIRST_WEEK_SYSTEMDEFAULT,
                                0
                            )?
                        );

                        println!("\n");
                        VariantClear(&mut value2)?;
                    } else {
                        break;
                    }
                }
                VariantClear(&mut value)?;
            } else {
                break;
            }
        }
        Ok(())
    }
}
