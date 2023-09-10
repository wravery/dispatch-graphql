use std::{
    env,
    fs::File,
    io::{self, Write},
    path::*,
    process::Command,
};

fn main() -> io::Result<()> {
    println!("cargo:rerun-if-changed=src/lib.rs");
    println!("cargo:rerun-if-changed=src/GraphQLService.idl");

    let mut idl_path = PathBuf::from("src");
    idl_path.push("GraphQLService.idl");
    let idl_path = idl_path.as_path().to_str().unwrap();

    let manifest_dir = PathBuf::from(env::var("CARGO_MANIFEST_DIR").unwrap());
    let out_dir = PathBuf::from(env::var("OUT_DIR").unwrap());
    let out_dir = out_dir.strip_prefix(manifest_dir).unwrap();
    let mut tlb_path = PathBuf::from(&out_dir);

    tlb_path.push("GraphQLService.tlb");
    let tlb_path = tlb_path.as_path().to_str().unwrap();
    let _ = Command::new("midl")
        .args(["/nologo", "/tlb", tlb_path, idl_path])
        .output()?;

    let mut rc_path = PathBuf::from(&out_dir);
    rc_path.push("GraphQLService.rc");
    let rc_path = rc_path.as_path().to_str().unwrap();
    let mut rc_file = File::create(rc_path)?;
    rc_file.write_all(
        b"1 typelib GraphQLService.tlb
",
    )?;

    let _ = Command::new("rc").args(["/nologo", rc_path]).output()?;

    let mut res_path = PathBuf::from(&out_dir);
    res_path.push("GraphQLService.res");
    let res_path = res_path.as_path().to_str().unwrap();

    println!("cargo:rustc-link-arg={res_path}");

    Ok(())
}
