plugins {
    id("com.android.library")
    id("org.jetbrains.kotlin.android")
}

android {
    namespace = "com.teamshryne.kinooffice"
    compileSdk = 35

    defaultConfig {
        minSdk = 26
        consumerProguardFiles("consumer-rules.pro")
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
    kotlinOptions { jvmTarget = "17" }
}

dependencies {
    implementation("org.apache.poi:poi-ooxml:5.3.0")
}

tasks.register<Copy>("copyRuntimeLibs") {
    from(configurations.releaseRuntimeClasspath)
    into(layout.buildDirectory.dir("runtime-libs"))
}
