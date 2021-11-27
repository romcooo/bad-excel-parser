import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

plugins {
    kotlin("jvm") version "1.6.0"
    `java-library`
    `maven-publish`
}

group = "net.romstu"
version = "1.0.7-SNAPSHOT"

val poi = "5.0.0"
val kotlinLogging = "2.1.0"
val logbackClassic = "1.2.7"

repositories {
    mavenCentral()
}

dependencies {
    testImplementation(kotlin("test"))

    implementation("org.apache.poi:poi:$poi")
    implementation("org.apache.poi:poi-ooxml:$poi")

    implementation("io.github.microutils:kotlin-logging:$kotlinLogging")
    implementation("ch.qos.logback:logback-classic:$logbackClassic")
}

tasks.test {
    useJUnitPlatform()
}

tasks.withType<KotlinCompile>() {
    kotlinOptions.jvmTarget = "11"
}

java {
    withSourcesJar()
    withJavadocJar()
}

project.afterEvaluate {
    publishing {
        publications {
            val mavenJava by creating(MavenPublication::class) {
                groupId = "net.romstu.exceltransformer"
                artifactId = "excel-transformer"
                version = "1.0.7"

                from(components["java"])
            }
        }
    }
}