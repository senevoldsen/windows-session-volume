use std::mem::MaybeUninit;
use std::ops::Deref;

use windows::runtime::{Interface, GUID};
use windows::Win32::Foundation::{PSTR, PWSTR};
use windows::Win32::Globalization::{WideCharToMultiByte, CP_UTF8};
use windows::Win32::Media::Audio::CoreAudio::{
    IAudioSessionControl, IAudioSessionEnumerator, IAudioSessionManager2, ISimpleAudioVolume,
};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::PropertiesSystem::PropVariantToStringAlloc;
use windows::Win32::{
    Media::Audio::CoreAudio::{
        eRender, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, MMDeviceEnumerator,
        DEVICE_STATE_ACTIVE,
    },
    System::{
        Com::{CoCreateInstance, CLSCTX_ALL},
        PropertiesSystem::IPropertyStore,
    },
};

const GUID_IAUDIO_SESSION_MANAGER2: &str = "77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F";
const PKEY_DEVICE_INTERFACE_FRIENDLY_NAME: &str = "a45c254e-df1c-4efd-8020-67d146a850e0";
const STGM_READ: u32 = 0x00000000;

struct ComMemPtr<T: AsComMemPtr> {
    value: T,
}

impl From<PWSTR> for ComMemPtr<PWSTR> {
    fn from(raw: PWSTR) -> Self {
        ComMemPtr { value: raw }
    }
}

impl<T> Deref for ComMemPtr<T>
where
    T: AsComMemPtr,
{
    type Target = T;

    fn deref(&self) -> &Self::Target {
        &self.value
    }
}

impl<T> Drop for ComMemPtr<T>
where
    T: AsComMemPtr,
{
    fn drop(&mut self) {
        unsafe {
            CoTaskMemFree(self.get_com_ptr());
        }
    }
}

trait AsComMemPtr {
    fn get_com_ptr(&self) -> *const ::std::ffi::c_void;
}

impl AsComMemPtr for PWSTR {
    fn get_com_ptr(&self) -> *const std::ffi::c_void {
        self.0 as *const _
    }
}

#[allow(dead_code)]
fn pwstr_to_string(input: PWSTR) -> String {
    // TODO: Seems like we keep a 0-terminator in the output
    // This causes a problem for direct comparisons.
    unsafe {
        let null_pstr: PSTR = PSTR(std::ptr::null_mut());
        let required_bytes: i32 = WideCharToMultiByte(
            CP_UTF8,
            0,
            input,
            -1,
            null_pstr,
            0,
            null_pstr,
            std::ptr::null_mut(),
        );
        assert!(required_bytes >= 0);
        let mut result = String::with_capacity(required_bytes as usize);
        let transformed_bytes = WideCharToMultiByte(
            CP_UTF8,
            0,
            input,
            -1,
            PSTR(result.as_mut_ptr()),
            required_bytes,
            null_pstr,
            std::ptr::null_mut(),
        );
        result.as_mut_vec().set_len(transformed_bytes as usize);

        assert_eq!(transformed_bytes, required_bytes);
        result
    }
}

#[allow(dead_code)]
fn pwstr_to_string2(input: PWSTR) -> String {
    unsafe {
        if input.0.is_null() {
            return String::new();
        }
        let mut end = input.0;
        while *end != 0 {
            end = end.add(1);
        }
        String::from_utf16_lossy(std::slice::from_raw_parts(
            input.0,
            end.offset_from(input.0) as _,
        ))
    }
}

struct AudioSessionIterator {
    count: i32,
    index: i32,
    enumerator: IAudioSessionEnumerator,
}

impl AudioSessionIterator {
    fn new(manager: &IAudioSessionManager2) -> Self {
        unsafe {
            let enumerator = manager.GetSessionEnumerator().unwrap();
            Self {
                count: enumerator.GetCount().unwrap(),
                index: 0,
                enumerator,
            }
        }
    }
}

impl Iterator for AudioSessionIterator {
    type Item = IAudioSessionControl;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index < self.count {
            self.index += 1;
            unsafe { Some(self.enumerator.GetSession(self.index - 1).unwrap()) }
        } else {
            None
        }
    }
}

pub fn find_application_for_device<Predicate>(
    device: &IMMDevice,
    pred: Predicate,
) -> Result<IAudioSessionControl, String>
where
    Predicate: Fn(&str) -> bool,
{
    unsafe {
        let session_manager = get_session_manager(device);
        let iterator = AudioSessionIterator::new(&session_manager);
        for session in iterator {
            let name: ComMemPtr<_> = session.GetDisplayName().unwrap().into();
            if pred(&pwstr_to_string2(*name)) {
                return Ok(session);
            }
        }
    }
    Err(String::from("audio session not found"))
}

pub fn applications_for_device(device: &IMMDevice) -> Vec<String> {
    let mut result = Vec::new();
    unsafe {
        let session_manager = get_session_manager(device);
        let iterator = AudioSessionIterator::new(&session_manager);
        for session in iterator {
            let name: ComMemPtr<_> = session.GetDisplayName().unwrap().into();
            result.push(pwstr_to_string2(*name));
        }
    }
    result
}

fn get_session_manager(device: &IMMDevice) -> IAudioSessionManager2 {
    let mut session_manager = MaybeUninit::<IAudioSessionManager2>::uninit();
    let guid: GUID = GUID_IAUDIO_SESSION_MANAGER2.into();
    unsafe {
        device
            .Activate(
                &guid,
                0u32,
                std::ptr::null(),
                std::ptr::addr_of_mut!(session_manager) as *mut *mut _,
            )
            .unwrap();
        session_manager.assume_init()
    }
}

fn propstore_find<Predicate>(prop_store: &IPropertyStore, pred: Predicate) -> bool
where
    Predicate: Fn(&str) -> bool,
{
    unsafe {
        let count = prop_store.GetCount().unwrap();

        for i in 0..count {
            let key = prop_store.GetAt(i).unwrap();
            if key.fmtid == PKEY_DEVICE_INTERFACE_FRIENDLY_NAME.into() {
                let prop_var = prop_store.GetValue(&key).unwrap();
                let raw_string_alloced: ComMemPtr<_> =
                    PropVariantToStringAlloc(&prop_var).unwrap().into();
                if pred(&pwstr_to_string2(*raw_string_alloced)) {
                    return true;
                }
            }
        }
    }
    false
}

pub fn find_device_with_friendly_name(arg_name: &str) -> Result<IMMDevice, String> {
    unsafe {
        let enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL).unwrap();

        let collection: IMMDeviceCollection = enumerator
            .EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)
            .unwrap();
        let num_devices = collection.GetCount().unwrap();
        for i in 0..num_devices {
            let device = collection.Item(i).unwrap();
            let prop_store = device.OpenPropertyStore(STGM_READ).unwrap();
            if propstore_find(&prop_store, |name| name.starts_with(arg_name)) {
                return Ok(device);
            }
        }
    };
    Err(String::from("no device with that name found"))
}

pub fn set_volume(session: &IAudioSessionControl, volume: f32) {
    let audio_control: ISimpleAudioVolume = session.cast().unwrap();
    unsafe {
        audio_control
            .SetMasterVolume(volume, &GUID::zeroed())
            .unwrap();
    }
}
