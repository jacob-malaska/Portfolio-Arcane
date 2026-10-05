// Start below-the-fold videos near the viewport, and pause them when offscreen.
// Keep source URLs in the HTML so videos still have normal browser controls.
document.addEventListener('DOMContentLoaded', function () {
    var videos = document.querySelectorAll('video[data-autoplay]');

    function play(video) {
        var result = video.play();
        if (result) result.catch(function () { video.controls = true; });
    }

    if (!('IntersectionObserver' in window)) {
        videos.forEach(function (video) { video.controls = true; });
        return;
    }

    var observer = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
            if (entry.isIntersecting) play(entry.target);
            else entry.target.pause();
        });
    }, { rootMargin: '200px 0px' });

    videos.forEach(function (video) { observer.observe(video); });
});
