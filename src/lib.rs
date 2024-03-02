use std::ops::Deref;

use windows::Win32::Devices::FunctionDiscovery::PKEY_DeviceInterface_FriendlyName;
use windows::Win32::System::Com::STGM_READ;
use windows::Win32::UI::Shell::PropertiesSystem::{IPropertyStore, PROPERTYKEY};
use windows::{
    core::{Interface, GUID, PWSTR},
    Win32::{
        Media::Audio::{
            eRender, IAudioSessionControl, IAudioSessionEnumerator, IAudioSessionManager2,
            IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, ISimpleAudioVolume,
            MMDeviceEnumerator, DEVICE_STATE_ACTIVE,
        },
        System::Com::{
            CoCreateInstance, CoTaskMemFree, StructuredStorage::PropVariantToStringAlloc,
            CLSCTX_ALL,
        },
    },
};

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
            CoTaskMemFree(Some(self.get_com_ptr()));
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
            if pred(&name.to_string().unwrap()) {
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
            result.push(name.to_string().unwrap());
        }
    }
    result
}

fn get_session_manager(device: &IMMDevice) -> IAudioSessionManager2 {
    unsafe {
        device
            .Activate(
                CLSCTX_ALL,
                None,
            )
            .unwrap()
    }
}

fn propstore_find<Predicate>(prop_store: &IPropertyStore, pred: Predicate) -> bool
where
    Predicate: Fn(&str) -> bool,
{
    unsafe {
        let count = prop_store.GetCount().unwrap();

        for i in 0..count {
            let mut key: PROPERTYKEY = Default::default();
            prop_store.GetAt(i, &mut key as *mut PROPERTYKEY).unwrap();
            if key.fmtid == PKEY_DeviceInterface_FriendlyName.fmtid {
                let prop_var = prop_store.GetValue(&key).unwrap();
                let raw_string_alloced: ComMemPtr<_> =
                PropVariantToStringAlloc(&prop_var).unwrap().into();
                if pred(&raw_string_alloced.to_string().unwrap()) {
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
