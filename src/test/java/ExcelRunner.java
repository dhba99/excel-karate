import com.intuit.karate.junit5.Karate;

class ExcelRunner {

    @Karate.Test
    Karate testSample() {
        return Karate.run("classpath:features/excel.feature").relativeTo(getClass());
    }

    @Karate.Test
    Karate testTags() {
        return Karate.run("tags").tags("@second").relativeTo(getClass());
    }

    @Karate.Test
    Karate testSystemProperty() {
        return Karate.run("classpath:karate/tags.feature")
                .tags("@second")
                .karateEnv("e2e")
                .systemProperty("foo", "bar");
    }

}